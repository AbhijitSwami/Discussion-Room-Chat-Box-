"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var sp_1 = require("@pnp/sp"); // Correct import
require("@pnp/sp/lists");
require("@pnp/sp/webs");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var RwBaseChatWebPart_module_scss_1 = __importDefault(require("./RwBaseChatWebPart.module.scss"));
var strings = __importStar(require("RwBaseChatWebPartStrings"));
var sp_http_1 = require("@microsoft/sp-http");
var SPOHelper_1 = require("./SPOHelper");
var listName;
var RwBaseChatWebPart = /** @class */ (function (_super) {
    __extends(RwBaseChatWebPart, _super);
    function RwBaseChatWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    RwBaseChatWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        // Initialize PnP SP object inside onInit
                        if (this.properties.autofetch) {
                            setInterval(function () { return _this.checkChatList(_this.properties.list); }, this.properties.interval);
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype.checkChatList = function (list) {
        return __awaiter(this, void 0, void 0, function () {
            var latestMessageTime, response, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        latestMessageTime = this._getLatestMessageTime();
                        return [4 /*yield*/, this._getNewMessages(latestMessageTime)];
                    case 1:
                        response = _a.sent();
                        if (response.value.length) {
                            this._renderList(response.value);
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        console.error("Error fetching messages:", error_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype._getLatestMessageTime = function () {
        var messagesContainer = this.domElement.querySelectorAll('.bubble');
        var lastMessageElement = messagesContainer[messagesContainer.length - 1];
        return lastMessageElement.dataset.timestamp || '';
    };
    RwBaseChatWebPart.prototype._getNewMessages = function (timestamp) {
        return __awaiter(this, void 0, void 0, function () {
            var filter, url, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        filter = timestamp ? "?$filter=Zeit gt '".concat(timestamp, "'") : '';
                        url = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/items").concat(filter);
                        return [4 /*yield*/, this.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, response.json()];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype.render = function () {
        var _this = this;
        listName = this.properties.list;
        this.domElement.innerHTML = "\n      <div class=\"".concat(RwBaseChatWebPart_module_scss_1.default.rwBaseChat, "\">\n        <div class=\"").concat(RwBaseChatWebPart_module_scss_1.default.container, "\">\n          <div class=\"").concat(RwBaseChatWebPart_module_scss_1.default.row, "\" style=\"background-color: ").concat(this.properties.background, ";\">\n            <span class=\"").concat(RwBaseChatWebPart_module_scss_1.default.title, "\">").concat((0, sp_lodash_subset_1.escape)(this.properties.title), "</span><br><br>\n            <div id=\"spListContainer\" style=\"max-height: 300px; overflow: auto; background-color: ").concat(this.properties.background, "\"></div>\n          </div>\n          <div class=\"").concat(RwBaseChatWebPart_module_scss_1.default.row, "\" style=\"background-color: ").concat(this.properties.background, ";\">\n            <table style=\"width: 100%; text-align: center;\">\n              <tr>\n                <td><button id=\"emoji\" class=\"").concat(RwBaseChatWebPart_module_scss_1.default.button2, "\"><span>&#128515;</span></button></td>\n                <td><textarea id=\"message\" placeholder=\"Your Message\" class=\"").concat(RwBaseChatWebPart_module_scss_1.default.textarea, "\"></textarea></td>\n                <td><button id=\"submit\" class=\"").concat(RwBaseChatWebPart_module_scss_1.default.button2, "\"><span>&#9993;</span></button></td>\n              </tr>\n            </table>\n            <div id=\"emojibar\" class=\"").concat(RwBaseChatWebPart_module_scss_1.default.emojibar, "\">\n              <button class=\"").concat(RwBaseChatWebPart_module_scss_1.default.emoji, "\">\uD83D\uDE02</button>\n              <button class=\"").concat(RwBaseChatWebPart_module_scss_1.default.emoji, "\">\uD83D\uDE05</button>\n              <button class=\"").concat(RwBaseChatWebPart_module_scss_1.default.emoji, "\">\uD83D\uDE09</button>\n              <button class=\"").concat(RwBaseChatWebPart_module_scss_1.default.emoji, "\">\uD83D\uDE07</button>\n              <button class=\"").concat(RwBaseChatWebPart_module_scss_1.default.emoji, "\">\uD83D\uDE0D</button>\n              <!-- More emojis -->\n            </div>\n          </div>\n        </div>\n      </div>");
        // Add event listeners
        var btnSubmit = this.domElement.querySelector("#submit");
        if (btnSubmit) { // Check if btnSubmit is not null
            btnSubmit.addEventListener("click", this.debounce(function () { return _this._getMessage(); }, 300));
        }
        var btnEmojiButton = this.domElement.querySelectorAll("." + RwBaseChatWebPart_module_scss_1.default.emoji);
        btnEmojiButton.forEach(function (item) {
            var emoji = item.textContent || ''; // Provide a fallback value if textContent is null
            item.addEventListener("click", function () { return _this.innerEmoji(emoji); });
        });
        var btnEmoji = this.domElement.querySelector("#emoji");
        if (btnEmoji) { // Check if btnEmoji is not null
            btnEmoji.addEventListener("click", function () {
                var emojiBar = _this.domElement.querySelector("#emojibar");
                if (emojiBar) { // Check if emojiBar is not null
                    emojiBar.classList.toggle(RwBaseChatWebPart_module_scss_1.default.emojibarOpen);
                    emojiBar.classList.toggle(RwBaseChatWebPart_module_scss_1.default.emojibar);
                }
            });
        }
        this.checkListCreation(this.properties.list);
        this._renderListAsync();
    };
    RwBaseChatWebPart.prototype.innerEmoji = function (emoji) {
        var txaMessage = this.domElement.querySelector('textarea'); // Cast to HTMLTextAreaElement
        if (txaMessage) { // Check if txaMessage is not null
            txaMessage.value += emoji; // Append the emoji to the textarea value
        }
    };
    RwBaseChatWebPart.prototype.debounce = function (fn, delay) {
        var timeoutID; // Use ReturnType to infer the type of timeoutID
        return function () {
            var _this = this;
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i] = arguments[_i];
            }
            if (timeoutID)
                clearTimeout(timeoutID);
            timeoutID = setTimeout(function () { return fn.apply(_this, args); }, delay);
        };
    };
    RwBaseChatWebPart.prototype._getMessage = function () {
        return __awaiter(this, void 0, void 0, function () {
            var user, messageElement, message, time, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        user = this.context.pageContext.user.displayName || this.context.pageContext.user.loginName;
                        messageElement = this.domElement.querySelector("textarea");
                        if (!messageElement) return [3 /*break*/, 3];
                        message = messageElement.value;
                        time = new Date().toLocaleString();
                        return [4 /*yield*/, this._addToList(time, user, message)];
                    case 1:
                        _a.sent();
                        messageElement.value = ''; // Clear the input
                        return [4 /*yield*/, this._renderListAsync()];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        console.warn("Textarea element not found.");
                        _a.label = 4;
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        error_2 = _a.sent();
                        console.error("Error submitting message:", error_2);
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype._addToList = function (zeit, user, message) {
        return __awaiter(this, void 0, void 0, function () {
            var error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, (0, SPOHelper_1.SPPost)({
                                url: "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/items"),
                                payload: { Zeit: zeit, User: user, Message: message }
                            })];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_3 = _a.sent();
                        console.error("Error adding message to list:", error_3);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype.checkListCreation = function (listName) {
        var _this = this;
        // Initialize PnP SP object in context
        var sp = (0, sp_1.spfi)().using((0, sp_1.SPFx)(this.context)); // Use spfi with SPFx context
        // Ensure list creation
        sp.web.lists.ensure(listName).then(function (_a) {
            var created = _a.created, list = _a.list;
            return __awaiter(_this, void 0, void 0, function () {
                var listDetails;
                return __generator(this, function (_b) {
                    switch (_b.label) {
                        case 0:
                            if (created) {
                                this._setFieldTypes(); // Call your method to set field types
                                console.log("List ".concat(listName, " was created."));
                            }
                            else {
                                console.log("List ".concat(listName, " already exists."));
                            }
                            return [4 /*yield*/, sp.web.lists.getByTitle(listName).select("Id", "Title")()];
                        case 1:
                            listDetails = _b.sent();
                            console.log("List Title: ".concat(listDetails.Title, ", List ID: ").concat(listDetails.Id));
                            return [2 /*return*/];
                    }
                });
            });
        }).catch(function (error) {
            console.error("Error creating or fetching list details:", error);
        });
    };
    RwBaseChatWebPart.prototype._setFieldTypes = function () {
        return __awaiter(this, void 0, void 0, function () {
            var baseUrl, fields, _i, fields_1, field, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        baseUrl = this.context.pageContext.web.absoluteUrl;
                        fields = ['Zeit', 'User', 'Message'];
                        _i = 0, fields_1 = fields;
                        _a.label = 1;
                    case 1:
                        if (!(_i < fields_1.length)) return [3 /*break*/, 4];
                        field = fields_1[_i];
                        return [4 /*yield*/, (0, SPOHelper_1.SPPost)({
                                url: "".concat(baseUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/fields"),
                                payload: { "FieldTypeKind": 2, "Title": field, "Required": true }
                            })];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        error_4 = _a.sent();
                        console.error("Error setting field types:", error_4);
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype._renderListAsync = function () {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._getListData()];
                    case 1:
                        response = _a.sent();
                        this._renderList(response.value);
                        return [3 /*break*/, 3];
                    case 2:
                        error_5 = _a.sent();
                        console.error("Error rendering list:", error_5);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    RwBaseChatWebPart.prototype._renderList = function (items) {
        var _this = this;
        var html = '';
        items.forEach(function (item) {
            var isCurrentUser = item.User === _this.context.pageContext.user.displayName;
            html += isCurrentUser ? "\n            <div class=\"".concat(RwBaseChatWebPart_module_scss_1.default.bubble, " ").concat(RwBaseChatWebPart_module_scss_1.default.alt, "\" id=\"msg_").concat(item.Id, "\">\n              <div class=\"").concat(RwBaseChatWebPart_module_scss_1.default.txt, "\">\n                <p class=\"").concat(RwBaseChatWebPart_module_scss_1.default.message, "\" style=\"text-align: right;\">").concat(item.Message, "</p>\n                <span class=\"").concat(RwBaseChatWebPart_module_scss_1.default.timestamp, "\"><button class=\"").concat(RwBaseChatWebPart_module_scss_1.default.answer, "\" id=\"msg_").concat(item.Id, "\">Answer</button> ").concat(item.Zeit, "</span>\n              </div>\n            </div>") : "\n            <div class=\"".concat(RwBaseChatWebPart_module_scss_1.default.bubble, "\">\n              <div class=\"").concat(RwBaseChatWebPart_module_scss_1.default.txt, "\">\n                <p class=\"").concat(RwBaseChatWebPart_module_scss_1.default.name, "\">").concat(item.User, "</p>\n                <p class=\"").concat(RwBaseChatWebPart_module_scss_1.default.message, "\">").concat(item.Message, "</p>\n                <span class=\"").concat(RwBaseChatWebPart_module_scss_1.default.timestamp, "\">").concat(item.Zeit, "</span>\n              </div>\n            </div>");
        });
        var listContainer = this.domElement.querySelector('#spListContainer');
        if (listContainer) { // Ensure listContainer is not null
            listContainer.innerHTML = html;
            listContainer.scrollTo(0, listContainer.scrollHeight);
        }
        else {
            console.error("List container element not found.");
        }
    };
    RwBaseChatWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get("".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists/getbytitle('").concat(listName, "')/items"), sp_http_1.SPHttpClient.configurations.v1).then(function (response) { return response.json(); });
    };
    Object.defineProperty(RwBaseChatWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    RwBaseChatWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: strings.PropertyPaneDescription },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                (0, sp_property_pane_1.PropertyPaneTextField)('title', { label: strings.TitleFieldLabel }),
                                (0, sp_property_pane_1.PropertyPaneTextField)('list', { label: strings.ListFieldLabel }),
                                (0, sp_property_pane_1.PropertyPaneCheckbox)('autofetch', { text: 'Automatically fetch new messages', checked: false }),
                                (0, sp_property_pane_1.PropertyPaneTextField)('interval', { label: strings.IntervalFieldLabel })
                            ]
                        },
                        {
                            groupName: strings.StyleGroupName,
                            groupFields: [
                                (0, sp_property_pane_1.PropertyPaneTextField)('background', { label: strings.BackgroundFieldLabel })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return RwBaseChatWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = RwBaseChatWebPart;
//# sourceMappingURL=RwBaseChatWebPart.js.map