var OfficeExtension;
(function (OfficeExtension) {
    var Action = (function () {
        function Action(actionInfo, isWriteOperation) {
            this.m_actionInfo = actionInfo;
            this.m_isWriteOperation = isWriteOperation;
        }
        Object.defineProperty(Action.prototype, "actionInfo", {
            get: function () {
                return this.m_actionInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Action.prototype, "isWriteOperation", {
            get: function () {
                return this.m_isWriteOperation;
            },
            enumerable: true,
            configurable: true
        });
        return Action;
    })();
    OfficeExtension.Action = Action;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ActionFactory = (function () {
        function ActionFactory() {
        }
        ActionFactory.createSetPropertyAction = function (context, parent, propertyName, value) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 4 /* SetProperty */,
                Name: propertyName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var args = [value];
            var referencedArgumentObjectPaths = OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var ret = new OfficeExtension.Action(actionInfo, true);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
            return ret;
        };
        ActionFactory.createMethodAction = function (context, parent, methodName, operationType, args) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 3 /* Method */,
                Name: methodName,
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var referencedArgumentObjectPaths = OfficeExtension.Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
            OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var isWriteOperation = operationType != 1 /* Read */;
            var ret = new OfficeExtension.Action(actionInfo, isWriteOperation);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
            return ret;
        };
        ActionFactory.createQueryAction = function (context, parent, queryOption) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 2 /* Query */,
                Name: "",
                ObjectPathId: parent._objectPath.objectPathInfo.Id,
            };
            actionInfo.QueryInfo = queryOption;
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(parent._objectPath);
            return ret;
        };
        ActionFactory.createInstantiateAction = function (context, obj) {
            OfficeExtension.Utility.validateObjectPath(obj);
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 1 /* Instantiate */,
                Name: "",
                ObjectPathId: obj._objectPath.objectPathInfo.Id
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addReferencedObjectPath(obj._objectPath);
            context._pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
            return ret;
        };
        ActionFactory.createTraceAction = function (context, message) {
            var actionInfo = {
                Id: context._nextId(),
                ActionType: 5 /* Trace */,
                Name: "Trace",
                ObjectPathId: 0
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context._pendingRequest.addAction(ret);
            context._pendingRequest.addTrace(actionInfo.Id, message);
            return ret;
        };
        return ActionFactory;
    })();
    OfficeExtension.ActionFactory = ActionFactory;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientObject = (function () {
        function ClientObject(context, objectPath) {
            OfficeExtension.Utility.checkArgumentNull(context, "context");
            this.m_context = context;
            this.m_objectPath = objectPath;
            if (this.m_objectPath) {
                if (!context._processingResult) {
                    OfficeExtension.ActionFactory.createInstantiateAction(context, this);
                    if ((context._autoCleanup) && (this._KeepReference)) {
                        context.references._autoAdd(this);
                    }
                }
            }
        }
        Object.defineProperty(ClientObject.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientObject.prototype, "_objectPath", {
            get: function () {
                return this.m_objectPath;
            },
            set: function (value) {
                this.m_objectPath = value;
            },
            enumerable: true,
            configurable: true
        });
        ClientObject.prototype._handleResult = function (value) {
        };
        return ClientObject;
    })();
    OfficeExtension.ClientObject = ClientObject;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientRequest = (function () {
        function ClientRequest(context) {
            this.m_context = context;
            this.m_actions = [];
            this.m_actionResultHandler = {};
            this.m_referencedObjectPaths = {};
            this.m_flags = 0 /* None */;
            this.m_traceInfos = {};
        }
        Object.defineProperty(ClientRequest.prototype, "flags", {
            get: function () {
                return this.m_flags;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequest.prototype, "traceInfos", {
            get: function () {
                return this.m_traceInfos;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype.addAction = function (action) {
            if (action.isWriteOperation) {
                this.m_flags = this.m_flags | 1 /* WriteOperation */;
            }
            this.m_actions.push(action);
        };
        Object.defineProperty(ClientRequest.prototype, "hasActions", {
            get: function () {
                return this.m_actions.length > 0;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequest.prototype.addTrace = function (actionId, message) {
            this.m_traceInfos[actionId] = message;
        };
        ClientRequest.prototype.addReferencedObjectPath = function (objectPath) {
            if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
                return;
            }
            if (!objectPath.isValid) {
                OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidObjectPath, OfficeExtension.Utility.getObjectPathExpression(objectPath));
            }
            while (objectPath) {
                if (objectPath.isWriteOperation) {
                    this.m_flags = this.m_flags | 1 /* WriteOperation */;
                }
                this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath;
                if (objectPath.objectPathInfo.ObjectPathType == 3 /* Method */) {
                    this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        ClientRequest.prototype.addReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.addReferencedObjectPath(objectPaths[i]);
                }
            }
        };
        ClientRequest.prototype.addActionResultHandler = function (action, resultHandler) {
            this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
        };
        ClientRequest.prototype.buildRequestMessageBody = function () {
            var objectPaths = {};
            for (var i in this.m_referencedObjectPaths) {
                objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
            }
            var actions = [];
            for (var index = 0; index < this.m_actions.length; index++) {
                actions.push(this.m_actions[index].actionInfo);
            }
            var ret = {
                Actions: actions,
                ObjectPaths: objectPaths
            };
            return ret;
        };
        ClientRequest.prototype.processResponse = function (msg) {
            if (msg && msg.Results) {
                for (var i = 0; i < msg.Results.length; i++) {
                    var actionResult = msg.Results[i];
                    var handler = this.m_actionResultHandler[actionResult.ActionId];
                    if (handler) {
                        handler._handleResult(actionResult.Value);
                    }
                }
            }
        };
        ClientRequest.prototype.invalidatePendingInvalidObjectPaths = function () {
            for (var i in this.m_referencedObjectPaths) {
                if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
                    this.m_referencedObjectPaths[i].isValid = false;
                }
            }
        };
        return ClientRequest;
    })();
    OfficeExtension.ClientRequest = ClientRequest;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientRequestContext = (function () {
        function ClientRequestContext(url) {
            this.m_nextId = 0;
            this.m_url = url;
            if (OfficeExtension.Utility.isNullOrEmptyString(this.m_url)) {
                this.m_url = OfficeExtension.Constants.localDocument;
            }
            this._processingResult = false;
            this._customData = OfficeExtension.Constants.iterativeExecutor;
            this._requestExecutor = new OfficeExtension.OfficeJsRequestExecutor();
            this.executeAsync = this.executeAsync.bind(this);
            this.sync = this.sync.bind(this);
        }
        Object.defineProperty(ClientRequestContext.prototype, "_pendingRequest", {
            get: function () {
                if (this.m_pendingRequest == null) {
                    this.m_pendingRequest = new OfficeExtension.ClientRequest(this);
                }
                return this.m_pendingRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "references", {
            get: function () {
                return this.trackedObjects;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ClientRequestContext.prototype, "trackedObjects", {
            get: function () {
                if (!this.m_references) {
                    this.m_references = new OfficeExtension.References(this);
                }
                return this.m_references;
            },
            enumerable: true,
            configurable: true
        });
        ClientRequestContext.prototype.load = function (clientObj, option) {
            OfficeExtension.Utility.validateContext(this, clientObj);
            var queryOption = {};
            if (typeof (option) == "string") {
                var select = option;
                queryOption.Select = this.parseSelectExpand(select);
            }
            else if (Array.isArray(option)) {
                queryOption.Select = option;
            }
            else if (typeof (option) == "object") {
                var loadOption = option;
                if (typeof (loadOption.select) == "string") {
                    queryOption.Select = this.parseSelectExpand(loadOption.select);
                }
                else if (Array.isArray(loadOption.select)) {
                    queryOption.Select = loadOption.select;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.select)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.select");
                }
                if (typeof (loadOption.expand) == "string") {
                    queryOption.Expand = this.parseSelectExpand(loadOption.expand);
                }
                else if (Array.isArray(loadOption.expand)) {
                    queryOption.Expand = loadOption.expand;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.expand)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.expand");
                }
                if (typeof (loadOption.top) == "number") {
                    queryOption.Top = loadOption.top;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.top)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.top");
                }
                if (typeof (loadOption.skip) == "number") {
                    queryOption.Skip = loadOption.skip;
                }
                else if (!OfficeExtension.Utility.isNullOrUndefined(loadOption.skip)) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option.skip");
                }
            }
            else if (!OfficeExtension.Utility.isNullOrUndefined(option)) {
                OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.invalidArgument, "option");
            }
            var action = OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
            this._pendingRequest.addActionResultHandler(action, clientObj);
        };
        ClientRequestContext.prototype.trace = function (message) {
            OfficeExtension.ActionFactory.createTraceAction(this, message);
        };
        ClientRequestContext.prototype.parseSelectExpand = function (select) {
            var args = [];
            if (!OfficeExtension.Utility.isNullOrEmptyString(select)) {
                var propertyNames = select.split(",");
                for (var i = 0; i < propertyNames.length; i++) {
                    var propertyName = propertyNames[i];
                    propertyName = propertyName.trim();
                    args.push(propertyName);
                }
            }
            return args;
        };
        ClientRequestContext.prototype.executeAsyncPrivate = function (doneCallback, failCallback) {
            var req = this._pendingRequest;
            if (!req.hasActions) {
                doneCallback();
                return;
            }
            this.m_pendingRequest = null;
            var msgBody = req.buildRequestMessageBody();
            var requestFlags = req.flags;
            var requestExecutor = this._requestExecutor;
            if (!requestExecutor) {
                requestExecutor = new OfficeExtension.OfficeJsRequestExecutor();
            }
            var requestExecutorRequestMessage = {
                Url: this.m_url,
                Headers: null,
                Body: msgBody
            };
            req.invalidatePendingInvalidObjectPaths();
            var thisObj = this;
            requestExecutor.executeAsync(this._customData, requestFlags, requestExecutorRequestMessage, function (response) {
                var error;
                var traceMessages = new Array();
                if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
                    error = new OfficeExtension.RuntimeError(response.ErrorCode, response.ErrorMessage, traceMessages, {});
                }
                else if (response.Body && response.Body.Error) {
                    error = new OfficeExtension.RuntimeError(response.Body.Error.Code, response.Body.Error.Message, traceMessages, {
                        errorLocation: response.Body.Error.Location
                    });
                }
                if (response.Body && response.Body.TraceIds) {
                    var traceMessageMap = req.traceInfos;
                    for (var i = 0; i < response.Body.TraceIds.length; i++) {
                        var traceId = response.Body.TraceIds[i];
                        var message = traceMessageMap[traceId];
                        traceMessages.push(message);
                    }
                }
                if (error) {
                    failCallback(error);
                    return;
                }
                else {
                    thisObj._processingResult = true;
                    try {
                        req.processResponse(response.Body);
                    }
                    finally {
                        thisObj._processingResult = false;
                    }
                    doneCallback();
                    return;
                }
            });
        };
        ClientRequestContext.prototype.executeAsync = function (passThroughValue) {
            return this.sync(passThroughValue);
        };
        ClientRequestContext.prototype.sync = function (passThroughValue) {
            var _this = this;
            OfficeExtension._EnsurePromise();
            return new OfficeExtension['Promise'](function (resolve, reject) {
                _this.executeAsyncPrivate(function () {
                    resolve(passThroughValue);
                }, function (error) {
                    reject(error);
                });
            });
        };
        ClientRequestContext._run = function (ctxInitializer, batch, numCleanupAttempts, retryDelay, onCleanupSuccess, onCleanupFailure) {
            if (numCleanupAttempts === void 0) { numCleanupAttempts = 3; }
            if (retryDelay === void 0) { retryDelay = 5000; }
            OfficeExtension._EnsurePromise();
            var starterPromise = new OfficeExtension['Promise'](function (resolve, reject) {
                resolve();
            });
            var ctx;
            var succeeded = false;
            var resultOrError;
            return starterPromise.then(function () {
                ctx = ctxInitializer();
                ctx._autoCleanup = true;
                var batchResult = batch(ctx);
                if (OfficeExtension.Utility.isNullOrUndefined(batchResult) || (typeof batchResult.then !== 'function')) {
                    OfficeExtension.Utility.throwError(OfficeExtension.ResourceStrings.runTaskAsyncMustReturnPromise);
                }
                return batchResult;
            }).then(function (batchResult) {
                return ctx.executeAsync(batchResult);
            }).then(function (result) {
                succeeded = true;
                resultOrError = result;
            }).catch(function (error) {
                resultOrError = error;
            }).then(function () {
                var itemsToRemove = ctx.references._retrieveAndClearAutoCleanupList();
                ctx._autoCleanup = false;
                for (var key in itemsToRemove) {
                    itemsToRemove[key]._objectPath.isValid = false;
                }
                var cleanupCounter = 0;
                attemptCleanup();
                function attemptCleanup() {
                    cleanupCounter++;
                    for (var key in itemsToRemove) {
                        ctx.references.remove(itemsToRemove[key]);
                    }
                    ctx.executeAsync().then(function () {
                        if (onCleanupSuccess) {
                            onCleanupSuccess(cleanupCounter);
                        }
                    }).catch(function () {
                        if (onCleanupFailure) {
                            onCleanupFailure(cleanupCounter);
                        }
                        if (cleanupCounter < numCleanupAttempts) {
                            setTimeout(function () {
                                attemptCleanup();
                            }, retryDelay);
                        }
                    });
                }
            }).then(function () {
                if (succeeded) {
                    return resultOrError;
                }
                else {
                    throw resultOrError;
                }
            });
        };
        ClientRequestContext.prototype._nextId = function () {
            return ++this.m_nextId;
        };
        return ClientRequestContext;
    })();
    OfficeExtension.ClientRequestContext = ClientRequestContext;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    (function (ClientRequestFlags) {
        ClientRequestFlags[ClientRequestFlags["None"] = 0] = "None";
        ClientRequestFlags[ClientRequestFlags["WriteOperation"] = 1] = "WriteOperation";
    })(OfficeExtension.ClientRequestFlags || (OfficeExtension.ClientRequestFlags = {}));
    var ClientRequestFlags = OfficeExtension.ClientRequestFlags;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientResult = (function () {
        function ClientResult() {
        }
        Object.defineProperty(ClientResult.prototype, "value", {
            get: function () {
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });
        ClientResult.prototype._handleResult = function (value) {
            this.m_value = value;
        };
        return ClientResult;
    })();
    OfficeExtension.ClientResult = ClientResult;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Constants = (function () {
        function Constants() {
        }
        Constants.getItemAt = "GetItemAt";
        Constants.id = "Id";
        Constants.idPrivate = "_Id";
        Constants.index = "_Index";
        Constants.items = "_Items";
        Constants.iterativeExecutor = "IterativeExecutor";
        Constants.localDocument = "http://document.localhost/";
        Constants.localDocumentApiPrefix = "http://document.localhost/_api/";
        Constants.referenceId = "_ReferenceId";
        return Constants;
    })();
    OfficeExtension.Constants = Constants;
})(OfficeExtension || (OfficeExtension = {}));
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var OfficeExtension;
(function (OfficeExtension) {
    var RuntimeError = (function (_super) {
        __extends(RuntimeError, _super);
        function RuntimeError(code, message, traceMessages, debugInfo) {
            _super.call(this, message);
            this.name = "OfficeExtension.RuntimeError";
            this.code = code;
            this.message = message;
            this.traceMessages = traceMessages;
            this.debugInfo = debugInfo;
        }
        RuntimeError.prototype.toString = function () {
            return this.code + ': ' + this.message;
        };
        return RuntimeError;
    })(Error);
    OfficeExtension.RuntimeError = RuntimeError;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ErrorCodes = (function () {
        function ErrorCodes() {
        }
        ErrorCodes.accessDenied = "AccessDenied";
        ErrorCodes.generalException = "GeneralException";
        ErrorCodes.activityLimitReached = "ActivityLimitReached";
        return ErrorCodes;
    })();
    OfficeExtension.ErrorCodes = ErrorCodes;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var InstantiateActionResultHandler = (function () {
        function InstantiateActionResultHandler(clientObject) {
            this.m_clientObject = clientObject;
        }
        InstantiateActionResultHandler.prototype._handleResult = function (value) {
            OfficeExtension.Utility.fixObjectPathIfNecessary(this.m_clientObject, value);
            if (value && !OfficeExtension.Utility.isNullOrUndefined(value[OfficeExtension.Constants.referenceId]) && this.m_clientObject._initReferenceId) {
                this.m_clientObject._initReferenceId(value[OfficeExtension.Constants.referenceId]);
            }
        };
        return InstantiateActionResultHandler;
    })();
    OfficeExtension.InstantiateActionResultHandler = InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    (function (RichApiRequestMessageIndex) {
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["CustomData"] = 0] = "CustomData";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["Method"] = 1] = "Method";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["PathAndQuery"] = 2] = "PathAndQuery";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["Headers"] = 3] = "Headers";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["Body"] = 4] = "Body";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["AppPermission"] = 5] = "AppPermission";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["RequestFlags"] = 6] = "RequestFlags";
    })(OfficeExtension.RichApiRequestMessageIndex || (OfficeExtension.RichApiRequestMessageIndex = {}));
    var RichApiRequestMessageIndex = OfficeExtension.RichApiRequestMessageIndex;
    (function (RichApiResponseMessageIndex) {
        RichApiResponseMessageIndex[RichApiResponseMessageIndex["StatusCode"] = 0] = "StatusCode";
        RichApiResponseMessageIndex[RichApiResponseMessageIndex["Headers"] = 1] = "Headers";
        RichApiResponseMessageIndex[RichApiResponseMessageIndex["Body"] = 2] = "Body";
    })(OfficeExtension.RichApiResponseMessageIndex || (OfficeExtension.RichApiResponseMessageIndex = {}));
    var RichApiResponseMessageIndex = OfficeExtension.RichApiResponseMessageIndex;
    ;
    (function (ActionType) {
        ActionType[ActionType["Instantiate"] = 1] = "Instantiate";
        ActionType[ActionType["Query"] = 2] = "Query";
        ActionType[ActionType["Method"] = 3] = "Method";
        ActionType[ActionType["SetProperty"] = 4] = "SetProperty";
        ActionType[ActionType["Trace"] = 5] = "Trace";
    })(OfficeExtension.ActionType || (OfficeExtension.ActionType = {}));
    var ActionType = OfficeExtension.ActionType;
    (function (ObjectPathType) {
        ObjectPathType[ObjectPathType["GlobalObject"] = 1] = "GlobalObject";
        ObjectPathType[ObjectPathType["NewObject"] = 2] = "NewObject";
        ObjectPathType[ObjectPathType["Method"] = 3] = "Method";
        ObjectPathType[ObjectPathType["Property"] = 4] = "Property";
        ObjectPathType[ObjectPathType["Indexer"] = 5] = "Indexer";
        ObjectPathType[ObjectPathType["ReferenceId"] = 6] = "ReferenceId";
    })(OfficeExtension.ObjectPathType || (OfficeExtension.ObjectPathType = {}));
    var ObjectPathType = OfficeExtension.ObjectPathType;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ObjectPath = (function () {
        function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
            this.m_objectPathInfo = objectPathInfo;
            this.m_parentObjectPath = parentObjectPath;
            this.m_isWriteOperation = false;
            this.m_isCollection = isCollection;
            this.m_isInvalidAfterRequest = isInvalidAfterRequest;
            this.m_isValid = true;
        }
        Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
            get: function () {
                return this.m_objectPathInfo;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
            get: function () {
                return this.m_isWriteOperation;
            },
            set: function (value) {
                this.m_isWriteOperation = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isCollection", {
            get: function () {
                return this.m_isCollection;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
            get: function () {
                return this.m_isInvalidAfterRequest;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
            get: function () {
                return this.m_parentObjectPath;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
            get: function () {
                return this.m_argumentObjectPaths;
            },
            set: function (value) {
                this.m_argumentObjectPaths = value;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ObjectPath.prototype, "isValid", {
            get: function () {
                return this.m_isValid;
            },
            set: function (value) {
                this.m_isValid = value;
            },
            enumerable: true,
            configurable: true
        });
        ObjectPath.prototype.updateUsingObjectData = function (value) {
            var referenceId = value[OfficeExtension.Constants.referenceId];
            if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                this.m_isInvalidAfterRequest = false;
                this.m_isValid = true;
                this.m_objectPathInfo.ObjectPathType = 6 /* ReferenceId */;
                this.m_objectPathInfo.Name = referenceId;
                this.m_objectPathInfo.ArgumentInfo = {};
                this.m_parentObjectPath = null;
                this.m_argumentObjectPaths = null;
                return;
            }
            if (this.parentObjectPath && this.parentObjectPath.isCollection) {
                var id = value[OfficeExtension.Constants.id];
                if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                    id = value[OfficeExtension.Constants.idPrivate];
                }
                if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
                    this.m_isInvalidAfterRequest = false;
                    this.m_isValid = true;
                    this.m_objectPathInfo.ObjectPathType = 5 /* Indexer */;
                    this.m_objectPathInfo.Name = "";
                    this.m_objectPathInfo.ArgumentInfo = {};
                    this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                    this.m_argumentObjectPaths = null;
                    return;
                }
            }
        };
        return ObjectPath;
    })();
    OfficeExtension.ObjectPath = ObjectPath;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ObjectPathFactory = (function () {
        function ObjectPathFactory() {
        }
        ObjectPathFactory.createGlobalObjectObjectPath = function (context) {
            var objectPathInfo = { Id: context._nextId(), ObjectPathType: 1 /* GlobalObject */, Name: "" };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
        };
        ObjectPathFactory.createNewObjectObjectPath = function (context, typeName, isCollection) {
            var objectPathInfo = { Id: context._nextId(), ObjectPathType: 2 /* NewObject */, Name: typeName };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
        };
        ObjectPathFactory.createPropertyObjectPath = function (context, parent, propertyName, isCollection, isInvalidAfterRequest) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 4 /* Property */,
                Name: propertyName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
            };
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
        };
        ObjectPathFactory.createIndexerObjectPath = function (context, parent, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5 /* Indexer */,
                Name: "",
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
        };
        ObjectPathFactory.createIndexerObjectPathUsingParentPath = function (context, parentObjectPath, args) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5 /* Indexer */,
                Name: "",
                ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
        };
        ObjectPathFactory.createMethodObjectPath = function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest) {
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3 /* Method */,
                Name: methodName,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = OfficeExtension.Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
            var ret = new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
            ret.argumentObjectPaths = argumentObjectPaths;
            ret.isWriteOperation = (operationType != 1 /* Read */);
            return ret;
        };
        ObjectPathFactory.createChildItemObjectPathUsingIndexer = function (context, parent, childItem) {
            var id = childItem[OfficeExtension.Constants.id];
            if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                id = childItem[OfficeExtension.Constants.idPrivate];
            }
            var objectPathInfo = objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 5 /* Indexer */,
                Name: "",
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [id];
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
        };
        ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function (context, parent, childItem, index) {
            var indexFromServer = childItem[OfficeExtension.Constants.index];
            if (indexFromServer) {
                index = indexFromServer;
            }
            var objectPathInfo = {
                Id: context._nextId(),
                ObjectPathType: 3 /* Method */,
                Name: OfficeExtension.Constants.getItemAt,
                ParentObjectPathId: parent._objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [index];
            return new OfficeExtension.ObjectPath(objectPathInfo, parent._objectPath, false, false);
        };
        ObjectPathFactory.createReferenceIdObjectPath = function (context, referenceId) {
            var objectPathInfo = { Id: context._nextId(), ObjectPathType: 6 /* ReferenceId */, Name: referenceId };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
        };
        return ObjectPathFactory;
    })();
    OfficeExtension.ObjectPathFactory = ObjectPathFactory;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var OfficeJsRequestExecutor = (function () {
        function OfficeJsRequestExecutor() {
        }
        OfficeJsRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage, callback) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            OfficeExtension.Utility.log("Request:");
            OfficeExtension.Utility.log(requestMessageText);
            var messageSafearray = OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", null, requestMessageText);
            OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
                OfficeExtension.Utility.log("Response:");
                OfficeExtension.Utility.log(JSON.stringify(result));
                var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
                if (result.status == "succeeded") {
                    var bodyText = OfficeExtension.RichApiMessageUtility.getResponseBody(result);
                    response.Body = JSON.parse(bodyText);
                    response.Headers = OfficeExtension.RichApiMessageUtility.getResponseHeaders(result);
                    callback(response);
                }
                else {
                    response.ErrorCode = OfficeExtension.ErrorCodes.generalException;
                    if (result.error.code == OfficeJsRequestExecutor.OfficeJsErrorCode_ooeNoCapability) {
                        response.ErrorCode = OfficeExtension.ErrorCodes.accessDenied;
                    }
                    else if (result.error.code == OfficeJsRequestExecutor.OfficeJsErrorCode_ooeActivityLimitReached) {
                        response.ErrorCode = OfficeExtension.ErrorCodes.activityLimitReached;
                    }
                    response.ErrorMessage = result.error.message;
                    callback(response);
                }
            });
        };
        OfficeJsRequestExecutor.OfficeJsErrorCode_ooeNoCapability = 7000;
        OfficeJsRequestExecutor.OfficeJsErrorCode_ooeActivityLimitReached = 5102;
        return OfficeJsRequestExecutor;
    })();
    OfficeExtension.OfficeJsRequestExecutor = OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var OfficeXHRSettings = (function () {
        function OfficeXHRSettings() {
        }
        return OfficeXHRSettings;
    })();
    OfficeExtension.OfficeXHRSettings = OfficeXHRSettings;
    function resetXHRFactory(oldFactory) {
        OfficeXHR.settings.oldxhr = oldFactory;
        return officeXHRFactory;
    }
    OfficeExtension.resetXHRFactory = resetXHRFactory;
    function officeXHRFactory() {
        return new OfficeXHR;
    }
    OfficeExtension.officeXHRFactory = officeXHRFactory;
    var OfficeXHR = (function () {
        function OfficeXHR() {
        }
        OfficeXHR.prototype.open = function (method, url) {
            this.m_method = method;
            this.m_url = url;
            if (this.m_url.toLowerCase().indexOf(OfficeExtension.Constants.localDocumentApiPrefix) == 0) {
                this.m_url = this.m_url.substr(OfficeExtension.Constants.localDocumentApiPrefix.length);
            }
            else {
                this.m_innerXhr = OfficeXHR.settings.oldxhr();
                var thisObj = this;
                this.m_innerXhr.onreadystatechange = function () {
                    thisObj.innerXhrOnreadystatechage();
                };
                this.m_innerXhr.open(method, this.m_url);
            }
        };
        OfficeXHR.prototype.abort = function () {
            if (this.m_innerXhr) {
                this.m_innerXhr.abort();
            }
        };
        OfficeXHR.prototype.send = function (body) {
            if (this.m_innerXhr) {
                this.m_innerXhr.send(body);
            }
            else {
                var thisObj = this;
                var requestFlags = 0 /* None */;
                if (!OfficeExtension.Utility.isReadonlyRestRequest(this.m_method)) {
                    requestFlags = 1 /* WriteOperation */;
                }
                var execFunction = OfficeXHR.settings.executeRichApiRequestAsync;
                if (!execFunction) {
                    execFunction = OSF.DDA.RichApi.executeRichApiRequestAsync;
                }
                execFunction(OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray('', requestFlags, this.m_method, this.m_url, this.m_requestHeaders, body), function (asyncResult) {
                    thisObj.officeContextRequestCallback(asyncResult);
                });
            }
        };
        OfficeXHR.prototype.setRequestHeader = function (header, value) {
            if (this.m_innerXhr) {
                this.m_innerXhr.setRequestHeader(header, value);
            }
            else {
                if (!this.m_requestHeaders) {
                    this.m_requestHeaders = {};
                }
                this.m_requestHeaders[header] = value;
            }
        };
        OfficeXHR.prototype.getResponseHeader = function (header) {
            if (this.m_responseHeaders) {
                return this.m_responseHeaders[header.toUpperCase()];
            }
            return null;
        };
        OfficeXHR.prototype.getAllResponseHeaders = function () {
            return this.m_allResponseHeaders;
        };
        OfficeXHR.prototype.overrideMimeType = function (mimeType) {
            if (this.m_innerXhr) {
                this.m_innerXhr.overrideMimeType(mimeType);
            }
        };
        OfficeXHR.prototype.innerXhrOnreadystatechage = function () {
            this.readyState = this.m_innerXhr.readyState;
            if (this.readyState == OfficeXHR.DONE) {
                this.status = this.m_innerXhr.status;
                this.statusText = this.m_innerXhr.statusText;
                this.responseText = this.m_innerXhr.responseText;
                this.response = this.m_innerXhr.response;
                this.responseType = this.m_innerXhr.responseType;
                this.setAllResponseHeaders(this.m_innerXhr.getAllResponseHeaders());
            }
            if (this.onreadystatechange) {
                this.onreadystatechange();
            }
        };
        OfficeXHR.prototype.officeContextRequestCallback = function (result) {
            this.readyState = OfficeXHR.DONE;
            if (result.status == "succeeded") {
                this.status = OfficeExtension.RichApiMessageUtility.getResponseStatusCode(result);
                this.m_responseHeaders = OfficeExtension.RichApiMessageUtility.getResponseHeaders(result);
                console.debug("ResponseHeaders=" + JSON.stringify(this.m_responseHeaders));
                this.responseText = OfficeExtension.RichApiMessageUtility.getResponseBody(result);
                console.debug("ResponseText=" + this.responseText);
                this.response = this.responseText;
            }
            else {
                this.status = 500;
                this.statusText = "Internal Error";
            }
            if (this.onreadystatechange) {
                this.onreadystatechange();
            }
        };
        OfficeXHR.prototype.setAllResponseHeaders = function (allResponseHeaders) {
            this.m_allResponseHeaders = allResponseHeaders;
            this.m_responseHeaders = {};
            if (this.m_allResponseHeaders != null) {
                var regex = new RegExp("\r?\n");
                var entries = this.m_allResponseHeaders.split(regex);
                for (var i = 0; i < entries.length; i++) {
                    var entry = entries[i];
                    if (entry != null) {
                        var index = entry.indexOf(':');
                        if (index > 0) {
                            var key = entry.substr(0, index);
                            var value = entry.substr(index + 1);
                            key = OfficeExtension.Utility.trim(key);
                            value = OfficeExtension.Utility.trim(value);
                            this.m_responseHeaders[key.toUpperCase()] = value;
                        }
                    }
                }
            }
        };
        OfficeXHR.UNSENT = 0;
        OfficeXHR.OPENED = 1;
        OfficeXHR.DONE = 4;
        OfficeXHR.settings = new OfficeXHRSettings();
        return OfficeXHR;
    })();
    OfficeExtension.OfficeXHR = OfficeXHR;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    function _EnsurePromise() {
        if (!OfficeExtension["Promise"]) {
            PromiseImpl.Init();
        }
    }
    OfficeExtension._EnsurePromise = _EnsurePromise;
    var PromiseImpl;
    (function (PromiseImpl) {
        function Init() {
            (function () {
                "use strict";
                function lib$es6$promise$utils$$objectOrFunction(x) {
                    return typeof x === 'function' || (typeof x === 'object' && x !== null);
                }
                function lib$es6$promise$utils$$isFunction(x) {
                    return typeof x === 'function';
                }
                function lib$es6$promise$utils$$isMaybeThenable(x) {
                    return typeof x === 'object' && x !== null;
                }
                var lib$es6$promise$utils$$_isArray;
                if (!Array.isArray) {
                    lib$es6$promise$utils$$_isArray = function (x) {
                        return Object.prototype.toString.call(x) === '[object Array]';
                    };
                }
                else {
                    lib$es6$promise$utils$$_isArray = Array.isArray;
                }
                var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                var lib$es6$promise$asap$$len = 0;
                var lib$es6$promise$asap$$toString = {}.toString;
                var lib$es6$promise$asap$$vertxNext;
                var lib$es6$promise$asap$$customSchedulerFn;
                var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                    lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                    lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                    lib$es6$promise$asap$$len += 2;
                    if (lib$es6$promise$asap$$len === 2) {
                        if (lib$es6$promise$asap$$customSchedulerFn) {
                            lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                        }
                        else {
                            lib$es6$promise$asap$$scheduleFlush();
                        }
                    }
                };
                function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                    lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                }
                function lib$es6$promise$asap$$setAsap(asapFn) {
                    lib$es6$promise$asap$$asap = asapFn;
                }
                var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' && typeof importScripts !== 'undefined' && typeof MessageChannel !== 'undefined';
                function lib$es6$promise$asap$$useNextTick() {
                    var nextTick = process.nextTick;
                    var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                    if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                        nextTick = setImmediate;
                    }
                    return function () {
                        nextTick(lib$es6$promise$asap$$flush);
                    };
                }
                function lib$es6$promise$asap$$useVertxTimer() {
                    return function () {
                        lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
                    };
                }
                function lib$es6$promise$asap$$useMutationObserver() {
                    var iterations = 0;
                    var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                    var node = document.createTextNode('');
                    observer.observe(node, { characterData: true });
                    return function () {
                        node.data = (iterations = ++iterations % 2);
                    };
                }
                function lib$es6$promise$asap$$useMessageChannel() {
                    var channel = new MessageChannel();
                    channel.port1.onmessage = lib$es6$promise$asap$$flush;
                    return function () {
                        channel.port2.postMessage(0);
                    };
                }
                function lib$es6$promise$asap$$useSetTimeout() {
                    return function () {
                        setTimeout(lib$es6$promise$asap$$flush, 1);
                    };
                }
                var lib$es6$promise$asap$$queue = new Array(1000);
                function lib$es6$promise$asap$$flush() {
                    for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                        var callback = lib$es6$promise$asap$$queue[i];
                        var arg = lib$es6$promise$asap$$queue[i + 1];
                        callback(arg);
                        lib$es6$promise$asap$$queue[i] = undefined;
                        lib$es6$promise$asap$$queue[i + 1] = undefined;
                    }
                    lib$es6$promise$asap$$len = 0;
                }
                function lib$es6$promise$asap$$attemptVertex() {
                    try {
                        var r = require;
                        var vertx = r('vertx');
                        lib$es6$promise$asap$$vertxNext = vertx.runOnLoop || vertx.runOnContext;
                        return lib$es6$promise$asap$$useVertxTimer();
                    }
                    catch (e) {
                        return lib$es6$promise$asap$$useSetTimeout();
                    }
                }
                var lib$es6$promise$asap$$scheduleFlush;
                if (lib$es6$promise$asap$$isNode) {
                    lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                }
                else if (lib$es6$promise$asap$$BrowserMutationObserver) {
                    lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMutationObserver();
                }
                else if (lib$es6$promise$asap$$isWorker) {
                    lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                }
                else if (lib$es6$promise$asap$$browserWindow === undefined && typeof require === 'function') {
                    lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$attemptVertex();
                }
                else {
                    lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                }
                function lib$es6$promise$$internal$$noop() {
                }
                var lib$es6$promise$$internal$$PENDING = void 0;
                var lib$es6$promise$$internal$$FULFILLED = 1;
                var lib$es6$promise$$internal$$REJECTED = 2;
                var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                function lib$es6$promise$$internal$$selfFullfillment() {
                    return new TypeError("You cannot resolve a promise with itself");
                }
                function lib$es6$promise$$internal$$cannotReturnOwn() {
                    return new TypeError('A promises callback cannot return that same promise.');
                }
                function lib$es6$promise$$internal$$getThen(promise) {
                    try {
                        return promise.then;
                    }
                    catch (error) {
                        lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
                        return lib$es6$promise$$internal$$GET_THEN_ERROR;
                    }
                }
                function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
                    try {
                        then.call(value, fulfillmentHandler, rejectionHandler);
                    }
                    catch (e) {
                        return e;
                    }
                }
                function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
                    lib$es6$promise$asap$$asap(function (promise) {
                        var sealed = false;
                        var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                            if (sealed) {
                                return;
                            }
                            sealed = true;
                            if (thenable !== value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }
                            else {
                                lib$es6$promise$$internal$$fulfill(promise, value);
                            }
                        }, function (reason) {
                            if (sealed) {
                                return;
                            }
                            sealed = true;
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }, 'Settle: ' + (promise._label || ' unknown promise'));
                        if (!sealed && error) {
                            sealed = true;
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                    }, promise);
                }
                function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                    if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                        lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                    }
                    else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
                        lib$es6$promise$$internal$$reject(promise, thenable._result);
                    }
                    else {
                        lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }, function (reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        });
                    }
                }
                function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
                    if (maybeThenable.constructor === promise.constructor) {
                        lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                    }
                    else {
                        var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                        if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                        }
                        else if (then === undefined) {
                            lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                        }
                        else if (lib$es6$promise$utils$$isFunction(then)) {
                            lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
                        }
                        else {
                            lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                        }
                    }
                }
                function lib$es6$promise$$internal$$resolve(promise, value) {
                    if (promise === value) {
                        lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
                    }
                    else if (lib$es6$promise$utils$$objectOrFunction(value)) {
                        lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
                    }
                    else {
                        lib$es6$promise$$internal$$fulfill(promise, value);
                    }
                }
                function lib$es6$promise$$internal$$publishRejection(promise) {
                    if (promise._onerror) {
                        promise._onerror(promise._result);
                    }
                    lib$es6$promise$$internal$$publish(promise);
                }
                function lib$es6$promise$$internal$$fulfill(promise, value) {
                    if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        return;
                    }
                    promise._result = value;
                    promise._state = lib$es6$promise$$internal$$FULFILLED;
                    if (promise._subscribers.length !== 0) {
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                    }
                }
                function lib$es6$promise$$internal$$reject(promise, reason) {
                    if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        return;
                    }
                    promise._state = lib$es6$promise$$internal$$REJECTED;
                    promise._result = reason;
                    lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                }
                function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                    var subscribers = parent._subscribers;
                    var length = subscribers.length;
                    parent._onerror = null;
                    subscribers[length] = child;
                    subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                    subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                    if (length === 0 && parent._state) {
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                    }
                }
                function lib$es6$promise$$internal$$publish(promise) {
                    var subscribers = promise._subscribers;
                    var settled = promise._state;
                    if (subscribers.length === 0) {
                        return;
                    }
                    var child, callback, detail = promise._result;
                    for (var i = 0; i < subscribers.length; i += 3) {
                        child = subscribers[i];
                        callback = subscribers[i + settled];
                        if (child) {
                            lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                        }
                        else {
                            callback(detail);
                        }
                    }
                    promise._subscribers.length = 0;
                }
                function lib$es6$promise$$internal$$ErrorObject() {
                    this.error = null;
                }
                var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                    try {
                        return callback(detail);
                    }
                    catch (e) {
                        lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                        return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                    }
                }
                function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                    var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                    if (hasCallback) {
                        value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                        if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                            failed = true;
                            error = value.error;
                            value = null;
                        }
                        else {
                            succeeded = true;
                        }
                        if (promise === value) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                            return;
                        }
                    }
                    else {
                        value = detail;
                        succeeded = true;
                    }
                    if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                    }
                    else if (hasCallback && succeeded) {
                        lib$es6$promise$$internal$$resolve(promise, value);
                    }
                    else if (failed) {
                        lib$es6$promise$$internal$$reject(promise, error);
                    }
                    else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                        lib$es6$promise$$internal$$fulfill(promise, value);
                    }
                    else if (settled === lib$es6$promise$$internal$$REJECTED) {
                        lib$es6$promise$$internal$$reject(promise, value);
                    }
                }
                function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
                    try {
                        resolver(function resolvePromise(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }, function rejectPromise(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        });
                    }
                    catch (e) {
                        lib$es6$promise$$internal$$reject(promise, e);
                    }
                }
                function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
                    var enumerator = this;
                    enumerator._instanceConstructor = Constructor;
                    enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                    if (enumerator._validateInput(input)) {
                        enumerator._input = input;
                        enumerator.length = input.length;
                        enumerator._remaining = input.length;
                        enumerator._init();
                        if (enumerator.length === 0) {
                            lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                        }
                        else {
                            enumerator.length = enumerator.length || 0;
                            enumerator._enumerate();
                            if (enumerator._remaining === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                        }
                    }
                    else {
                        lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                    }
                }
                lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                    return lib$es6$promise$utils$$isArray(input);
                };
                lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                    return new OfficeExtension.Error('Array Methods must be provided an Array');
                };
                lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                    this._result = new Array(this.length);
                };
                var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                    var enumerator = this;
                    var length = enumerator.length;
                    var promise = enumerator.promise;
                    var input = enumerator._input;
                    for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                        enumerator._eachEntry(input[i], i);
                    }
                };
                lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                    var enumerator = this;
                    var c = enumerator._instanceConstructor;
                    if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                        if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                            entry._onerror = null;
                            enumerator._settledAt(entry._state, i, entry._result);
                        }
                        else {
                            enumerator._willSettleAt(c.resolve(entry), i);
                        }
                    }
                    else {
                        enumerator._remaining--;
                        enumerator._result[i] = entry;
                    }
                };
                lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                    var enumerator = this;
                    var promise = enumerator.promise;
                    if (promise._state === lib$es6$promise$$internal$$PENDING) {
                        enumerator._remaining--;
                        if (state === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, value);
                        }
                        else {
                            enumerator._result[i] = value;
                        }
                    }
                    if (enumerator._remaining === 0) {
                        lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                    }
                };
                lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                    var enumerator = this;
                    lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                        enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                    }, function (reason) {
                        enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                    });
                };
                function lib$es6$promise$promise$all$$all(entries) {
                    return new lib$es6$promise$enumerator$$default(this, entries).promise;
                }
                var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                function lib$es6$promise$promise$race$$race(entries) {
                    var Constructor = this;
                    var promise = new Constructor(lib$es6$promise$$internal$$noop);
                    if (!lib$es6$promise$utils$$isArray(entries)) {
                        lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                        return promise;
                    }
                    var length = entries.length;
                    function onFulfillment(value) {
                        lib$es6$promise$$internal$$resolve(promise, value);
                    }
                    function onRejection(reason) {
                        lib$es6$promise$$internal$$reject(promise, reason);
                    }
                    for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                        lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                    }
                    return promise;
                }
                var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                function lib$es6$promise$promise$resolve$$resolve(object) {
                    var Constructor = this;
                    if (object && typeof object === 'object' && object.constructor === Constructor) {
                        return object;
                    }
                    var promise = new Constructor(lib$es6$promise$$internal$$noop);
                    lib$es6$promise$$internal$$resolve(promise, object);
                    return promise;
                }
                var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                function lib$es6$promise$promise$reject$$reject(reason) {
                    var Constructor = this;
                    var promise = new Constructor(lib$es6$promise$$internal$$noop);
                    lib$es6$promise$$internal$$reject(promise, reason);
                    return promise;
                }
                var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                var lib$es6$promise$promise$$counter = 0;
                function lib$es6$promise$promise$$needsResolver() {
                    throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                }
                function lib$es6$promise$promise$$needsNew() {
                    throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                }
                var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                function lib$es6$promise$promise$$Promise(resolver) {
                    this._id = lib$es6$promise$promise$$counter++;
                    this._state = undefined;
                    this._result = undefined;
                    this._subscribers = [];
                    if (lib$es6$promise$$internal$$noop !== resolver) {
                        if (!lib$es6$promise$utils$$isFunction(resolver)) {
                            lib$es6$promise$promise$$needsResolver();
                        }
                        if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                            lib$es6$promise$promise$$needsNew();
                        }
                        lib$es6$promise$$internal$$initializePromise(this, resolver);
                    }
                }
                lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                lib$es6$promise$promise$$Promise.prototype = {
                    constructor: lib$es6$promise$promise$$Promise,
                    then: function (onFulfillment, onRejection) {
                        var parent = this;
                        var state = parent._state;
                        if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                            return this;
                        }
                        var child = new this.constructor(lib$es6$promise$$internal$$noop);
                        var result = parent._result;
                        if (state) {
                            var callback = arguments[state - 1];
                            lib$es6$promise$asap$$asap(function () {
                                lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
                            });
                        }
                        else {
                            lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
                        }
                        return child;
                    },
                    'catch': function (onRejection) {
                        return this.then(null, onRejection);
                    }
                };
                OfficeExtension["Promise"] = lib$es6$promise$promise$$default;
            }).call(this);
        }
        PromiseImpl.Init = Init;
    })(PromiseImpl || (PromiseImpl = {}));
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    (function (OperationType) {
        OperationType[OperationType["Default"] = 0] = "Default";
        OperationType[OperationType["Read"] = 1] = "Read";
    })(OfficeExtension.OperationType || (OfficeExtension.OperationType = {}));
    var OperationType = OfficeExtension.OperationType;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var References = (function () {
        function References(context) {
            this._autoCleanupList = {};
            this.m_context = context;
        }
        References.prototype.add = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._addCommon(item, true); });
            }
            else {
                this._addCommon(param, true);
            }
        };
        References.prototype._autoAdd = function (object) {
            this._addCommon(object, false);
            this._autoCleanupList[object._objectPath.objectPathInfo.Id] = object;
        };
        References.prototype._addCommon = function (object, isExplicitlyAdded) {
            var referenceId = object[OfficeExtension.Constants.referenceId];
            if (OfficeExtension.Utility.isNullOrEmptyString(referenceId) && object._KeepReference) {
                object._KeepReference();
                OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, object);
                if (isExplicitlyAdded && this.m_context._autoCleanup) {
                    delete this._autoCleanupList[object._objectPath.objectPathInfo.Id];
                }
            }
        };
        References.prototype.remove = function (param) {
            var _this = this;
            if (Array.isArray(param)) {
                param.forEach(function (item) { return _this._removeCommon(item); });
            }
            else {
                this._removeCommon(param);
            }
        };
        References.prototype._removeCommon = function (object) {
            var referenceId = object[OfficeExtension.Constants.referenceId];
            if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                var rootObject = this.m_context._rootObject;
                if (rootObject._RemoveReference) {
                    rootObject._RemoveReference(referenceId);
                }
            }
        };
        References.prototype._retrieveAndClearAutoCleanupList = function () {
            var list = this._autoCleanupList;
            this._autoCleanupList = {};
            return list;
        };
        return References;
    })();
    OfficeExtension.References = References;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ResourceStrings = (function () {
        function ResourceStrings() {
        }
        ResourceStrings.invalidObjectPath = "InvalidObjectPath";
        ResourceStrings.propertyNotLoaded = "PropertyNotLoaded";
        ResourceStrings.invalidRequestContext = "InvalidRequestContext";
        ResourceStrings.invalidArgument = "InvalidArgument";
        ResourceStrings.runTaskAsyncMustReturnPromise = "RunTaskAsyncMustReturnPromise";
        return ResourceStrings;
    })();
    OfficeExtension.ResourceStrings = ResourceStrings;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var RichApiMessageUtility = (function () {
        function RichApiMessageUtility() {
        }
        RichApiMessageUtility.buildRequestMessageSafeArray = function (customData, requestFlags, method, path, headers, body) {
            var headerArray = [];
            if (headers) {
                for (var headerName in headers) {
                    headerArray.push(headerName);
                    headerArray.push(headers[headerName]);
                }
            }
            var appPermission = 0;
            var solutionId = "";
            var instanceId = "";
            var marketplaceType = "";
            return [
                customData,
                method,
                path,
                headerArray,
                body,
                appPermission,
                requestFlags,
                solutionId,
                instanceId,
                marketplaceType
            ];
        };
        RichApiMessageUtility.getResponseBody = function (result) {
            return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseHeaders = function (result) {
            return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseBodyFromSafeArray = function (data) {
            var ret = data[2 /* Body */];
            if (typeof (ret) === "string") {
                return ret;
            }
            var arr = ret;
            return arr.join("");
        };
        RichApiMessageUtility.getResponseHeadersFromSafeArray = function (data) {
            var arrayHeader = data[1 /* Headers */];
            if (!arrayHeader) {
                return null;
            }
            var headers = {};
            for (var i = 0; i < arrayHeader.length - 1; i += 2) {
                headers[arrayHeader[i]] = arrayHeader[i + 1];
            }
            return headers;
        };
        RichApiMessageUtility.getResponseStatusCode = function (result) {
            return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
        };
        RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function (data) {
            return data[0 /* StatusCode */];
        };
        return RichApiMessageUtility;
    })();
    OfficeExtension.RichApiMessageUtility = RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Utility = (function () {
        function Utility() {
        }
        Utility.checkArgumentNull = function (value, name) {
        };
        Utility.isNullOrUndefined = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof (value) === "undefined") {
                return true;
            }
            return false;
        };
        Utility.isUndefined = function (value) {
            if (typeof (value) === "undefined") {
                return true;
            }
            return false;
        };
        Utility.isNullOrEmptyString = function (value) {
            if (value === null) {
                return true;
            }
            if (typeof (value) === "undefined") {
                return true;
            }
            if (value.length == 0) {
                return true;
            }
            return false;
        };
        Utility.trim = function (str) {
            return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
        };
        Utility.caseInsensitiveCompareString = function (str1, str2) {
            if (Utility.isNullOrUndefined(str1)) {
                return Utility.isNullOrUndefined(str2);
            }
            else {
                if (Utility.isNullOrUndefined(str2)) {
                    return false;
                }
                else {
                    return str1.toUpperCase() == str2.toUpperCase();
                }
            }
        };
        Utility.isReadonlyRestRequest = function (method) {
            return Utility.caseInsensitiveCompareString(method, "GET");
        };
        Utility.setMethodArguments = function (context, argumentInfo, args) {
            if (Utility.isNullOrUndefined(args)) {
                return null;
            }
            var referencedObjectPaths = new Array();
            var referencedObjectPathIds = new Array();
            var hasOne = false;
            for (var i = 0; i < args.length; i++) {
                if (args[i] instanceof OfficeExtension.ClientObject) {
                    var clientObject = args[i];
                    Utility.validateContext(context, clientObject);
                    args[i] = clientObject._objectPath.objectPathInfo.Id;
                    referencedObjectPathIds.push(clientObject._objectPath.objectPathInfo.Id);
                    referencedObjectPaths.push(clientObject._objectPath);
                    hasOne = true;
                }
                else {
                    referencedObjectPathIds.push(0);
                }
            }
            argumentInfo.Arguments = args;
            if (hasOne) {
                argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds;
                return referencedObjectPaths;
            }
            return null;
        };
        Utility.fixObjectPathIfNecessary = function (clientObject, value) {
            if (clientObject && clientObject._objectPath && value) {
                clientObject._objectPath.updateUsingObjectData(value);
            }
        };
        Utility.validateObjectPath = function (clientObject) {
            var objectPath = clientObject._objectPath;
            while (objectPath) {
                if (!objectPath.isValid) {
                    var pathExpression = Utility.getObjectPathExpression(objectPath);
                    Utility.throwError(OfficeExtension.ResourceStrings.invalidObjectPath, pathExpression);
                }
                objectPath = objectPath.parentObjectPath;
            }
        };
        Utility.validateReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    var objectPath = objectPaths[i];
                    while (objectPath) {
                        if (!objectPath.isValid) {
                            var pathExpression = Utility.getObjectPathExpression(objectPath);
                            Utility.throwError(OfficeExtension.ResourceStrings.invalidObjectPath, pathExpression);
                        }
                        objectPath = objectPath.parentObjectPath;
                    }
                }
            }
        };
        Utility.validateContext = function (context, obj) {
            if (obj && obj.context !== context) {
                Utility.throwError(OfficeExtension.ResourceStrings.invalidRequestContext);
            }
        };
        Utility.log = function (message) {
            if (Utility._logEnabled && window.console && window.console.log) {
                window.console.log(message);
            }
        };
        Utility.load = function (clientObj, option) {
            clientObj.context.load(clientObj, option);
        };
        Utility.throwError = function (resourceId, arg) {
            throw new OfficeExtension.RuntimeError(resourceId, Utility._getResourceString(resourceId, arg), new Array(), {});
        };
        Utility.createRuntimeError = function (code, message, location) {
            return new OfficeExtension.RuntimeError(code, message, [], { errorLocation: location });
        };
        Utility._getResourceString = function (resourceId, arg) {
            var ret = resourceId;
            if (window.Strings && window.Strings.OfficeOM) {
                var stringName = "L_" + resourceId;
                var stringValue = window.Strings.OfficeOM[stringName];
                if (stringValue) {
                    ret = stringValue;
                }
            }
            if (!Utility.isNullOrUndefined(arg)) {
                ret = ret.replace("{0}", arg);
            }
            return ret;
        };
        Utility.throwIfNotLoaded = function (propertyName, fieldValue) {
            if (Utility.isUndefined(fieldValue) && propertyName.charCodeAt(0) != Utility.s_underscoreCharCode) {
                Utility.throwError(OfficeExtension.ResourceStrings.propertyNotLoaded, propertyName);
            }
        };
        Utility.getObjectPathExpression = function (objectPath) {
            var ret = "";
            while (objectPath) {
                switch (objectPath.objectPathInfo.ObjectPathType) {
                    case 1 /* GlobalObject */:
                        ret = ret;
                        break;
                    case 2 /* NewObject */:
                        ret = "new()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 3 /* Method */:
                        ret = Utility.normalizeName(objectPath.objectPathInfo.Name) + "()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 4 /* Property */:
                        ret = Utility.normalizeName(objectPath.objectPathInfo.Name) + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 5 /* Indexer */:
                        ret = "getItem()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                    case 6 /* ReferenceId */:
                        ret = "_reference()" + (ret.length > 0 ? "." : "") + ret;
                        break;
                }
                objectPath = objectPath.parentObjectPath;
            }
            return ret;
        };
        Utility._createPromiseFromResult = function (value) {
            OfficeExtension._EnsurePromise();
            return new OfficeExtension['Promise'](function (resolve, reject) {
                resolve(value);
            });
        };
        Utility._addActionResultHandler = function (clientObj, action, resultHandler) {
            clientObj.context._pendingRequest.addActionResultHandler(action, resultHandler);
        };
        Utility._handleNavigationPropertyResults = function (clientObj, objectValue, propertyNames) {
            for (var i = 0; i < propertyNames.length - 1; i += 2) {
                if (!Utility.isUndefined(objectValue[propertyNames[i + 1]])) {
                    clientObj[propertyNames[i]]._handleResult(objectValue[propertyNames[i + 1]]);
                }
            }
        };
        Utility.normalizeName = function (name) {
            return name.substr(0, 1).toLowerCase() + name.substr(1);
        };
        Utility._logEnabled = false;
        Utility.s_underscoreCharCode = "_".charCodeAt(0);
        return Utility;
    })();
    OfficeExtension.Utility = Utility;
})(OfficeExtension || (OfficeExtension = {}));
