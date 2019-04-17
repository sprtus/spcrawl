"use strict";
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
var sp_pnp_node_1 = require("sp-pnp-node");
var pnpjs_1 = require("@pnp/pnpjs");
// Migrate ACS policy pages
var Policies = /** @class */ (function () {
    function Policies() {
    }
    Policies.run = function (config) {
        return __awaiter(this, void 0, void 0, function () {
            var _loop_1, _i, _a, siteUrl;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _loop_1 = function (siteUrl) {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        console.log("Connecting to " + siteUrl);
                                        // Connect
                                        return [4 /*yield*/, new sp_pnp_node_1.PnpNode().init().then(function (settings) { return __awaiter(_this, void 0, void 0, function () {
                                                var web, pages, _i, _a, page, relatedPolicies, fileCheckedIn, fileApproved;
                                                return __generator(this, function (_b) {
                                                    switch (_b.label) {
                                                        case 0:
                                                            console.log("  Connected to " + settings.siteUrl);
                                                            web = new pnpjs_1.Web(siteUrl + "/policy-center");
                                                            return [4 /*yield*/, web.lists.getByTitle('Pages').items.select('*').expand('file').getAll().catch(function (e) { return console.error; })];
                                                        case 1:
                                                            pages = _b.sent();
                                                            _i = 0, _a = pages;
                                                            _b.label = 2;
                                                        case 2:
                                                            if (!(_i < _a.length)) return [3 /*break*/, 8];
                                                            page = _a[_i];
                                                            if (!page.File || !page.File.ServerRelativeUrl)
                                                                return [3 /*break*/, 7];
                                                            console.log("Updating " + page.File.ServerRelativeUrl);
                                                            // Check out
                                                            return [4 /*yield*/, web.getFileByServerRelativeUrl(page.File.ServerRelativeUrl).checkout().catch(function (e) { })];
                                                        case 3:
                                                            // Check out
                                                            _b.sent();
                                                            console.log('  Checked out');
                                                            relatedPolicies = page.ACSRelatedPolicies.replace(/\/policies-forms\/policies-procedures\//ig, '/policy-center/');
                                                            // Replace related IDs
                                                            // for (const rPage of pages as any[]) {
                                                            //   const searchTitle = new RegExp(`"Id":"\\d+","Title":"${rPage.Title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'ig');
                                                            //   relatedPolicies = relatedPolicies.replace(searchTitle, `"Id":"${rPage.ID}","Title":"${rPage.Title}"`);
                                                            // }
                                                            // Update related policies
                                                            return [4 /*yield*/, web.lists.getByTitle('Pages').items.getById(page.ID).update({
                                                                    ACSRelatedPolicies: relatedPolicies,
                                                                }).catch(function (e) { return console.error; })];
                                                        case 4:
                                                            // Replace related IDs
                                                            // for (const rPage of pages as any[]) {
                                                            //   const searchTitle = new RegExp(`"Id":"\\d+","Title":"${rPage.Title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'ig');
                                                            //   relatedPolicies = relatedPolicies.replace(searchTitle, `"Id":"${rPage.ID}","Title":"${rPage.Title}"`);
                                                            // }
                                                            // Update related policies
                                                            _b.sent();
                                                            console.log('  Updated');
                                                            return [4 /*yield*/, web.getFileByServerRelativeUrl(page.File.ServerRelativeUrl).checkin('Checked in via automated migration', pnpjs_1.CheckinType.Major).catch(function (e) { })];
                                                        case 5:
                                                            fileCheckedIn = _b.sent();
                                                            console.log('  Checked in');
                                                            return [4 /*yield*/, web.getFileByServerRelativeUrl(page.File.ServerRelativeUrl).approve('Approved via automated migration').catch(function (e) { })];
                                                        case 6:
                                                            fileApproved = _b.sent();
                                                            console.log('  Approved');
                                                            _b.label = 7;
                                                        case 7:
                                                            _i++;
                                                            return [3 /*break*/, 2];
                                                        case 8: return [2 /*return*/];
                                                    }
                                                });
                                            }); }).catch(console.error)];
                                    case 1:
                                        // Connect
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        };
                        _i = 0, _a = config.sites;
                        _b.label = 1;
                    case 1:
                        if (!(_i < _a.length)) return [3 /*break*/, 4];
                        siteUrl = _a[_i];
                        return [5 /*yield**/, _loop_1(siteUrl)];
                    case 2:
                        _b.sent();
                        _b.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4:
                        console.log('Done.');
                        return [2 /*return*/];
                }
            });
        });
    };
    return Policies;
}());
exports.Policies = Policies;