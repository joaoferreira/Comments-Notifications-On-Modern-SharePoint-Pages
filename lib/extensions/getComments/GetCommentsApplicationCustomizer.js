var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
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
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
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
import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { SPHttpClient } from '@microsoft/sp-http';
import pnp from "sp-pnp-js";
import { PermissionKind } from "@pnp/sp";
var LOG_SOURCE = 'GetCommentsApplicationCustomizer';
/** A Custom Action which can be run during execution of a Client Side Application */
var GetCommentsApplicationCustomizer = (function (_super) {
    __extends(GetCommentsApplicationCustomizer, _super);
    function GetCommentsApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    GetCommentsApplicationCustomizer.prototype.onInit = function () {
        var _this = this;
        pnp.setup({
            spfxContext: this.context
        });
        pnp.sp.web.currentUserHasPermissions(PermissionKind.ManageWeb).then(function (perms) {
            if (perms) {
                _this._renderPlaceHolders();
            }
        });
        //this.getSitePageComments();
        //this.getPages();
        return Promise.resolve();
    };
    GetCommentsApplicationCustomizer.prototype.getSitePageComments = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            var currentWebUrl, response, responseJSON, commentsNumber, _i, _a, entry;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        currentWebUrl = this.context.pageContext.web.serverRelativeUrl;
                        return [4 /*yield*/, this.context.spHttpClient.get(currentWebUrl + "/_api/web/lists/GetByTitle('Site Pages')/GetItemById(" + id + ")/Comments?$expand=replies,likedBy,replies/likedBy&$top=10&$inlineCount=AllPages", SPHttpClient.configurations.v1)];
                    case 1:
                        response = _b.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        responseJSON = _b.sent();
                        commentsNumber = 0;
                        for (_i = 0, _a = responseJSON.value; _i < _a.length; _i++) {
                            entry = _a[_i];
                            commentsNumber++;
                            commentsNumber = commentsNumber + entry.replyCount;
                        }
                        this.UpdateItem(id, commentsNumber);
                        console.log('comments: ' + commentsNumber + 'page id: ' + id);
                        return [2 /*return*/];
                }
            });
        });
    };
    GetCommentsApplicationCustomizer.prototype.getPages = function () {
        return __awaiter(this, void 0, void 0, function () {
            var currentWebUrl, response, responseJSON, pageId, _i, _a, entry;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        currentWebUrl = this.context.pageContext.web.serverRelativeUrl;
                        return [4 /*yield*/, this.context.spHttpClient.get(currentWebUrl + "/_api/web/lists/GetByTitle('Site Pages')/items", SPHttpClient.configurations.v1)];
                    case 1:
                        response = _b.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        responseJSON = _b.sent();
                        for (_i = 0, _a = responseJSON.value; _i < _a.length; _i++) {
                            entry = _a[_i];
                            pageId = entry.ID;
                            this.getSitePageComments(pageId);
                        }
                        alert('SharePoint is processing the page comments!');
                        return [2 /*return*/];
                }
            });
        });
    };
    GetCommentsApplicationCustomizer.prototype.UpdateItem = function (id, comments) {
        pnp.sp.web.lists.getByTitle("Site Pages").items.getById(id).update({
            Previous_x0020_Comments: comments
        }).then(console.log)
            .catch(console.log);
    };
    GetCommentsApplicationCustomizer.prototype._renderPlaceHolders = function () {
        // Handling the bottom placeholder
        if (!this._bottomPlaceholder) {
            this._bottomPlaceholder =
                this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, { onDispose: this._onDispose });
            // The extension should not assume that the expected placeholder is available.
            if (!this._bottomPlaceholder) {
                console.error('The expected placeholder (Bottom) was not found.');
                return;
            }
            if (this.properties) {
                if (this._bottomPlaceholder.domElement) {
                    this._bottomPlaceholder.domElement.innerHTML = "\n              <div id=\"checkComments\" Title=\"Get Comments Notfications\" style=\"position: absolute; bottom: 0; width: 22px; height: 18px; left: 10px; z-index: 100; padding: 10px; cursor: pointer;\" class=\"ms-bgColor-themeDark ms-fontColor-white \">\n                <i class=\"ms-Icon ms-Icon--Message\" aria-hidden=\"true\" style=\"font-size: 20px;\"></i>\n              </div>";
                }
            }
        }
        var ctx = this;
        document.getElementById('checkComments').onclick = function () { ctx.getPages(); };
    };
    GetCommentsApplicationCustomizer.prototype._onDispose = function () {
        console.log('Disposed Coments.');
    };
    __decorate([
        override
    ], GetCommentsApplicationCustomizer.prototype, "onInit", null);
    return GetCommentsApplicationCustomizer;
}(BaseApplicationCustomizer));
export default GetCommentsApplicationCustomizer;

//# sourceMappingURL=GetCommentsApplicationCustomizer.js.map
