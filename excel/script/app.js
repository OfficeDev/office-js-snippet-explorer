var excelSamplesApp = angular.module("excelSamplesApp", ['ngRoute']);
var insideOffice = false;
var consoleErrorFunction;

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

excelSamplesApp.config(['$routeProvider', function ($routeProvider) {
	$routeProvider
		.when('/samples',
			{
				controller: 'SamplesController',
				templateUrl: 'partials/samples.html'
			})
		.when('/testAll',
			{
				controller: 'TestAllController',
				templateUrl: 'partials/testAll.html'
			})
		.otherwise({redirectTo: '/samples' });
}]);

excelSamplesApp.factory("excelSamplesFactory", ['$http', function ($http) {
	var factory = {};
	
	factory.getSamples = function() {
		return $http.get('samples/samples.json');
	};

	factory.getSampleCode = function(filename) {
		return $http.get('samples/' + filename);
	};

	return factory;
}]);

excelSamplesApp.controller("SamplesController", function($scope, excelSamplesFactory) {
	$scope.samples = [{ name: "Loading..." }];
	$scope.selectedSample = { description: "No sample loaded" };
	$scope.insideOffice = insideOffice;
	
	MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
			"/excel/script/EditorIntelliSense/ExcelLatest.txt",
			"/excel/script/EditorIntelliSense/Office.Runtime.txt",
			"/excel/script/EditorIntelliSense/Helpers.txt",
			"/excel/script/EditorIntelliSense/jquery.txt",
		]);
	
	MonacoEditorIntegration.setDirty = function() {
		if ($scope.selectedSample.code) {
			$scope.selectedSample = { description: $scope.selectedSample.description + " (modified)" };
			$scope.$apply();
		}
	}
	
	excelSamplesFactory.getSamples().then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		appInsights.trackEvent("SampleLoaded", {name:$scope.selectedSample.name});
		excelSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
			$scope.selectedSample.code = addErrorHandlingIfNeeded(response.data);
			$scope.insideOffice = insideOffice;
			MonacoEditorIntegration.setJavaScriptText($scope.selectedSample.code);
			MonacoEditorIntegration.resizeEditor();
		});
	};
	
	$scope.runSelectedSample = function() {
		var script = MonacoEditorIntegration.getJavaScriptToRun().replace(/console.log/g, "logComment");
		try {
			eval(script);
		} catch (e) {
			logComment(e.name + ": " + e.message);
		}
	}

});


excelSamplesApp.controller("TestAllController", function($scope, $q, excelSamplesFactory) {
	$scope.insideOffice = insideOffice;

	excelSamplesFactory.getSamples().then(function (response) {
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
				excelSamplesFactory.getSampleCode(sample.filename).then(function (response) {
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
	return sampleCode.replace("ctx.executeAsync().then();", "ctx.executeAsync().then(function() {\r\n    console.log(\"done\");\r\n}, function(error) {\r\n    console.log(\"An error occurred: \" + error.errorCode + \":\" + error.errorMessage);\r\n});");
}

function addErrorHandlingIfNeeded(sampleCode) {
	if (!insideOffice) return sampleCode;
	return addErrorHandling(sampleCode);	
}