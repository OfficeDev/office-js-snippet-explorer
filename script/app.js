var officeJsSnippetApp = angular.module("officeJsSnippetApp", ['ngRoute']);
var insideOffice = false;
var consoleErrorFunction;

var rootUrl = document.location;

var logComment = function(message) {
	var consoleElement;
	consoleElement = document.getElementById('console');
	consoleElement.innerHTML += message + '\n';
	consoleElement.scrollTop = consoleElement.scrollHeight;
}

Office.initialize = function (reason) {
	insideOffice = true;	
	console.log('Add-in initialized, redirecting console.log() to console textArea');
	consoleErrorFunction = console.error;
	console.error = logComment;
	// get Angular scope from the known DOM element
    var e = document.getElementById('samplesContainer');
    var scope = angular.element(e).scope();
    // update the model with a wrap in $apply(fn) which will refresh the view for us
    scope.$apply(function() {
        scope.insideOffice = true;
    }); 
};

officeJsSnippetApp.config(['$routeProvider', function ($routeProvider) {
	$routeProvider
		.when('/snippets/:app',
			{
				controller: 'SamplesController',
				templateUrl: 'partials/snippet-browser.html'
			})
		.when('/add-in/:app',
			{
				controller: 'SamplesController',
				templateUrl: 'partials/add-in.html'
			})
		.when('/testAll',
			{
				controller: 'TestAllController',
				templateUrl: 'partials/testAll.html'
			})
		.otherwise({redirectTo: '/add-in/excel' });
}]);

officeJsSnippetApp.factory("snippetFactory", ['$http', function ($http) {
	var factory = {};
	
	factory.getSamples = function(app) {
		return $http.get(app + '-snippets/samples.json');
	};

	factory.getSampleCode = function(app, filename) {
		return $http.get(app + '-snippets/' + filename);
	};

	return factory;
}]);

officeJsSnippetApp.controller("SamplesController", function($scope, $routeParams, snippetFactory) {
	$scope.samples = [{ name: "Loading..." }];
	$scope.selectedSample = { description: "No snippet loaded" };
	$scope.insideOffice = insideOffice;
	
	CodeEditorIntegration.initializeJsEditor('TxtRichApiScript', [
			"/editorIntelliSense/ExcelLatest.txt",
			"/editorIntelliSense/WordLatest.txt",
			"/editorIntelliSense/OfficeCommon.txt",
			"/editorIntelliSense/OfficeDocument.txt"
	]);
	
	CodeEditorIntegration.setDirty = function() {
		if ($scope.selectedSample.code) {
			$scope.selectedSample = { description: $scope.selectedSample.description + " (modified)" };
			$scope.$apply();
		}
	}
	
	snippetFactory.getSamples($routeParams["app"]).then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		appInsights.trackEvent("SampleLoaded", {name:$scope.selectedSample.name});
		snippetFactory.getSampleCode($routeParams["app"], $scope.selectedSample.filename).then(function (response) {
            $scope.selectedSample.code = addErrorHandlingIfNeeded(response.data);
			$scope.insideOffice = insideOffice;
			CodeEditorIntegration.setJavaScriptText($scope.selectedSample.code);
			CodeEditorIntegration.resizeEditor();
		});
	};
	
	$scope.runSelectedSample = function() {
		var script = CodeEditorIntegration.getJavaScriptToRun().replace(/console.log/g, "logComment");
		
		if (isTrulyJavaScript(script)) {
			try {
				eval(script);
			} catch (e) {
				logComment(e.name + ": " + e.message);
			}	
		} else {
			CodeEditorIntegration.getEditorTextAsJavaScript().then(function (output) {
				if (output == null) {
					logComment("Invalid JavaScript / TypeScript. Please fix the errors shown in the code editor and try again.");
				} else {
					eval(output.content);
				}
			});
		}	
	}
});

officeJsSnippetApp.controller("TestAllController", function($scope, $q, snippetFactory) {
	$scope.insideOffice = insideOffice;

	snippetFactory.getSamples().then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		appInsights.trackEvent("SampleLoaded", {name:$scope.selectedSample.name});

	};

	$scope.runSamples = function() {
		
		var promiseProducingSampleFunctions = new Array();
		
		for (var i = 1; i < $scope.samples.length; i++) {
			promiseProducingSampleFunctions.push(createRunSample(i));
		}
		
		var result = createRunSample(0);
		result = result();
		promiseProducingSampleFunctions.forEach(function (f) {
			result = result.then(f);
		});
		
		function createRunSample(sampleIndex) {
			
			var sample = $scope.samples[sampleIndex];
			
			return function() {
				var deferred = $q.defer();
				//logComment("running next call");
				sample.runStatus = "Loading";
				snippetFactory.getSampleCode(sample.filename).then(function (response) {
					sample.code = addTestResults(addDeferredErrorHandling(response.data)).replace(/console.log/g, "logComment");
					sample.runStatus = "Running";
					try {
						//logComment(sample.code);
						eval(sample.code);
					} catch (e) {
						sample.runStatus = "Error: " + e.name + ": " + e.message;
						deferred.resolve();
					}
				});
				
				return deferred.promise;
			}
		}
	}
	
	$scope.refreshResults = function() {
		$scope.$apply();
	}

});

function addTestResults(sampleCode) {
	return sampleCode.replace("console.log(\"done\");", "sample.runStatus = \"Success\"; deferred.resolve();");
}

function addDeferredErrorHandling(sampleCode) {
	return sampleCode.replace("ctx.executeAsync().then();", "ctx.executeAsync().then(function() {\r\n    console.log(\"done\");\r\n}, function(error) {\r\n    sample.runStatus = \"Error: \" + error.errorCode + \":\" + error.errorMessage; deferred.resolve(); });");
}

function addErrorHandling(sampleCode) {
	return sampleCode.replace("\r\n}).catch(function (error) {\r\n	console.log(error);\r\n});", "\r\n}).catch(function(error) {\r\n    console.log(\"Error: \" + error);\r\n    if (error instanceof OfficeExtension.Error) {\r\n        console.log(\"Debug info: \" + JSON.stringify(error.debugInfo));\r\n    }\r\n});");
}

function addErrorHandlingIfNeeded(sampleCode) {
	if (!insideOffice) return sampleCode;
	return addErrorHandling(sampleCode);	
}

/** returns whether the text is truly javascript (as opposed to typescript) */
function isTrulyJavaScript(text) {
	try {
		new Function(text);
		return true;
	} catch (syntaxError) {
		return false;
	}
}

