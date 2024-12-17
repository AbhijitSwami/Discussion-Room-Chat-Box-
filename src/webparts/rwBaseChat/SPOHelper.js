"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.SPDelete = exports.SPUpdate = exports.SPMultiFileUpload = exports.SPFileUpload = exports.SPPost = exports.SPGet = exports.GetDigest = void 0;
/**
 * Checks if an object is of JSON type.
 */
var isObject = function (obj) { return obj !== undefined && obj !== null && obj.constructor === Object; };
/**
 * Base HTTP client wrapper around fetch.
 */
var BaseClient = function (url, options) {
    if (options === void 0) { options = {}; }
    return __awaiter(void 0, void 0, void 0, function () {
        var error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, fetch(url, options)];
                case 1: return [2 /*return*/, _a.sent()];
                case 2:
                    error_1 = _a.sent();
                    console.error("Error in BaseClient:", error_1);
                    throw error_1;
                case 3: return [2 /*return*/];
            }
        });
    });
};
/**
 * Fetches JSON response using BaseClient.
 */
var GetJson = function (url, options) {
    if (options === void 0) { options = {}; }
    return __awaiter(void 0, void 0, void 0, function () {
        var response;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, BaseClient(url, options)];
                case 1:
                    response = _a.sent();
                    return [2 /*return*/, response.json()];
            }
        });
    });
};
var JSFY = JSON.stringify;
var _headers = {
    credentials: "include",
    Accept: "application/json; odata=nometadata",
    "Content-Type": "application/json; odata=nometadata"
};
var defaultDigest = { FormDigestValue: "" };
var copyObj = function (obj) { return (__assign({}, obj)); };
var MergeObj = function (obj1, obj2) { return (__assign(__assign({}, obj1), obj2)); };
var _Post = function (_a) {
    var url = _a.url, _b = _a.payload, payload = _b === void 0 ? {} : _b, _c = _a.hdrs, hdrs = _c === void 0 ? {} : _c, _d = _a.isBlobOrArrayBuffer, isBlobOrArrayBuffer = _d === void 0 ? false : _d;
    return __awaiter(void 0, void 0, void 0, function () {
        var headers, body, response;
        return __generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    headers = MergeObj(_headers, hdrs);
                    body = isBlobOrArrayBuffer ? payload : JSFY(payload);
                    return [4 /*yield*/, BaseClient(url, { method: "POST", body: body, headers: headers })];
                case 1:
                    response = _e.sent();
                    if (headers["IF-MATCH"])
                        return [2 /*return*/, response];
                    return [2 /*return*/, response.json()];
            }
        });
    });
};
/**
 * Retrieves the Request Digest.
 */
var GetDigest = function (url) { return __awaiter(void 0, void 0, void 0, function () {
    var urlSegments;
    return __generator(this, function (_a) {
        if (!url || typeof url !== "string")
            throw new Error("Invalid URL");
        urlSegments = url.toLowerCase().split("/_api");
        if (urlSegments.length > 1)
            url = "".concat(urlSegments[0], "/_api/contextinfo");
        return [2 /*return*/, _Post({ url: url })];
    });
}); };
exports.GetDigest = GetDigest;
/**
 * Get request for SharePoint API.
 */
var SPGet = function (url) { return __awaiter(void 0, void 0, void 0, function () {
    var headers;
    return __generator(this, function (_a) {
        headers = copyObj(_headers);
        return [2 /*return*/, GetJson(url, { headers: headers })];
    });
}); };
exports.SPGet = SPGet;
/**
 * Post request with Request Digest.
 */
var SPPost = function (_a) {
    var url = _a.url, _b = _a.payload, payload = _b === void 0 ? {} : _b, _c = _a.hdrs, hdrs = _c === void 0 ? {} : _c, _d = _a.digest, digest = _d === void 0 ? defaultDigest : _d, _e = _a.isBlobOrArrayBuffer, isBlobOrArrayBuffer = _e === void 0 ? false : _e;
    return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_f) {
            switch (_f.label) {
                case 0:
                    if (!(!digest || (isObject(digest) && !digest.FormDigestValue))) return [3 /*break*/, 2];
                    return [4 /*yield*/, (0, exports.GetDigest)(url)];
                case 1:
                    digest = _f.sent();
                    _f.label = 2;
                case 2:
                    hdrs["X-RequestDigest"] = digest.FormDigestValue;
                    return [2 /*return*/, _Post({ url: url, hdrs: hdrs, payload: payload, isBlobOrArrayBuffer: isBlobOrArrayBuffer })];
            }
        });
    });
};
exports.SPPost = SPPost;
/**
 * Uploads a file as an attachment.
 */
var SPFileUpload = function (_a) {
    var url = _a.url, payload = _a.payload, _b = _a.hdrs, hdrs = _b === void 0 ? {} : _b, _c = _a.digest, digest = _c === void 0 ? defaultDigest : _c;
    return __awaiter(void 0, void 0, void 0, function () {
        return __generator(this, function (_d) {
            return [2 /*return*/, (0, exports.SPPost)({ isBlobOrArrayBuffer: true, url: url, hdrs: hdrs, digest: digest, payload: payload })];
        });
    });
};
exports.SPFileUpload = SPFileUpload;
/**
 * Uploads multiple files to SharePoint.
 */
var SPMultiFileUpload = function (_a) {
    var url = _a.url, _b = _a.hdrs, hdrs = _b === void 0 ? {} : _b, _c = _a.files, files = _c === void 0 ? [] : _c, _d = _a.digest, digest = _d === void 0 ? defaultDigest : _d;
    return __awaiter(void 0, void 0, void 0, function () {
        var promises;
        return __generator(this, function (_e) {
            promises = files.map(function (file) {
                return (0, exports.SPFileUpload)({ url: "".concat(url, "add(FileName='").concat(file.fileName, "',overwrite='true')"), payload: file.data, hdrs: hdrs, digest: digest });
            });
            return [2 /*return*/, Promise.all(promises)];
        });
    });
};
exports.SPMultiFileUpload = SPMultiFileUpload;
/**
 * Update request with Request Digest.
 */
var SPUpdate = function (_a) {
    var url = _a.url, _b = _a.payload, payload = _b === void 0 ? {} : _b, _c = _a.digest, digest = _c === void 0 ? defaultDigest : _c;
    return __awaiter(void 0, void 0, void 0, function () {
        var hdrs;
        return __generator(this, function (_d) {
            hdrs = { "IF-MATCH": "*", "X-HTTP-Method": "MERGE" };
            return [2 /*return*/, (0, exports.SPPost)({ url: url, hdrs: hdrs, payload: payload, digest: digest })];
        });
    });
};
exports.SPUpdate = SPUpdate;
/**
 * Delete request with Request Digest.
 */
var SPDelete = function (url, digest) {
    if (digest === void 0) { digest = defaultDigest; }
    return __awaiter(void 0, void 0, void 0, function () {
        var hdrs;
        return __generator(this, function (_a) {
            hdrs = { "IF-MATCH": "*", "X-HTTP-Method": "DELETE" };
            return [2 /*return*/, (0, exports.SPPost)({ url: url, hdrs: hdrs, digest: digest })];
        });
    });
};
exports.SPDelete = SPDelete;
/**
 * Generates a unique file name with current timestamp.
 */
// const GetUniqueFileName = (): string => `${new Date().getTime()}.txt`;
 //sourceMappingURL=SPOHelper.js.map