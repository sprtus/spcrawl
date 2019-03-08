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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var lodash_1 = __importDefault(require("lodash"));
var IContentType_1 = require("./data/IContentType");
var IField_1 = require("./data/IField");
var IWeb_1 = require("./data/IWeb");
var sp_pnp_node_1 = require("sp-pnp-node");
var Utility_1 = require("./Utility");
var pnpjs_1 = require("@pnp/pnpjs");
var chalk_1 = __importDefault(require("chalk"));
var fs_1 = __importDefault(require("fs"));
// Scan a site collection
var Scan = /** @class */ (function () {
    function Scan() {
    }
    Scan.run = function (config) {
        return __awaiter(this, void 0, void 0, function () {
            var sites, counters, _loop_1, _i, _a, siteUrl;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        sites = [];
                        counters = {
                            contentTypes: 0,
                            fields: 0,
                            sites: config.sites.length,
                            webs: 0,
                        };
                        _loop_1 = function (siteUrl) {
                            var site;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        console.log("Crawling webs in " + siteUrl);
                                        site = {
                                            LastScan: new Date().toISOString(),
                                            Url: siteUrl,
                                            Webs: [],
                                        };
                                        // Connect
                                        return [4 /*yield*/, new sp_pnp_node_1.PnpNode().init().then(function (settings) { return __awaiter(_this, void 0, void 0, function () {
                                                var _a, _b, webUrls, webUrl, web, webResponse, webData, _i, _c, prop, ctypes, _d, ctypes_1, ctype, ctypeData, _e, _f, prop, fields, _g, fields_1, field, fieldData, _h, _j, prop, subwebs, _k, subwebs_1, subweb;
                                                return __generator(this, function (_l) {
                                                    switch (_l.label) {
                                                        case 0:
                                                            webUrls = [siteUrl];
                                                            _l.label = 1;
                                                        case 1:
                                                            if (!webUrls.length) return [3 /*break*/, 6];
                                                            webUrl = lodash_1.default.trim(webUrls.shift(), '/');
                                                            console.log("  - " + chalk_1.default.whiteBright.bold(webUrl));
                                                            web = new pnpjs_1.Web(webUrl);
                                                            return [4 /*yield*/, web.select.apply(web, Object.keys(IWeb_1.NewWeb)).get().catch(console.error)];
                                                        case 2:
                                                            webResponse = _l.sent();
                                                            webData = lodash_1.default.cloneDeep(IWeb_1.NewWeb);
                                                            for (_i = 0, _c = Object.keys(IWeb_1.NewWeb); _i < _c.length; _i++) {
                                                                prop = _c[_i];
                                                                if (webResponse[prop] !== undefined) {
                                                                    webData[prop] = webResponse[prop];
                                                                }
                                                            }
                                                            site.Webs.push(webData);
                                                            return [4 /*yield*/, (_a = web.contentTypes.filter("Group ne '_Hidden'")).select.apply(_a, Object.keys(IContentType_1.NewContentType)).get().catch(console.error)];
                                                        case 3:
                                                            ctypes = _l.sent();
                                                            for (_d = 0, ctypes_1 = ctypes; _d < ctypes_1.length; _d++) {
                                                                ctype = ctypes_1[_d];
                                                                ctypeData = lodash_1.default.clone(IContentType_1.NewContentType);
                                                                for (_e = 0, _f = Object.keys(IContentType_1.NewContentType); _e < _f.length; _e++) {
                                                                    prop = _f[_e];
                                                                    if (ctype[prop] !== undefined) {
                                                                        if (ctype[prop] && ctype[prop].StringValue)
                                                                            ctypeData[prop] = ctype[prop].StringValue;
                                                                        else
                                                                            ctypeData[prop] = ctype[prop];
                                                                    }
                                                                }
                                                                webData.ContentTypes.push(ctypeData);
                                                                // Increment content type counter
                                                                counters.contentTypes++;
                                                            }
                                                            return [4 /*yield*/, (_b = web.fields.filter("Hidden eq false")).select.apply(_b, Object.keys(IField_1.NewField)).get().catch(console.error)];
                                                        case 4:
                                                            fields = _l.sent();
                                                            for (_g = 0, fields_1 = fields; _g < fields_1.length; _g++) {
                                                                field = fields_1[_g];
                                                                fieldData = lodash_1.default.clone(IField_1.NewField);
                                                                for (_h = 0, _j = Object.keys(IField_1.NewField); _h < _j.length; _h++) {
                                                                    prop = _j[_h];
                                                                    if (field[prop] !== undefined) {
                                                                        if (field[prop] && field[prop].StringValue)
                                                                            fieldData[prop] = field[prop].StringValue;
                                                                        else
                                                                            fieldData[prop] = field[prop];
                                                                    }
                                                                }
                                                                webData.Fields.push(fieldData);
                                                                // Increment field counter
                                                                counters.fields++;
                                                            }
                                                            return [4 /*yield*/, web.webs.get().catch(console.error)];
                                                        case 5:
                                                            subwebs = _l.sent();
                                                            // Add subwebs to webs array
                                                            for (_k = 0, subwebs_1 = subwebs; _k < subwebs_1.length; _k++) {
                                                                subweb = subwebs_1[_k];
                                                                webUrls.push(subweb.Url);
                                                            }
                                                            // Increment webs counter
                                                            counters.webs++;
                                                            return [3 /*break*/, 1];
                                                        case 6: return [2 /*return*/];
                                                    }
                                                });
                                            }); }).catch(console.error)];
                                    case 1:
                                        // Connect
                                        _a.sent();
                                        // Add site
                                        sites.push(site);
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
                        // Write to file
                        console.log("Found " + counters.webs + " webs, " + counters.contentTypes + " content types, and " + counters.fields + " fields in " + counters.sites + " sites.\n  " + chalk_1.default.greenBright(Utility_1.Utility.path('webs.json')));
                        try {
                            fs_1.default.writeFileSync(Utility_1.Utility.path('webs.json'), JSON.stringify(sites, null, 2), { flag: 'w' });
                        }
                        catch (e) {
                            console.error(e);
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    return Scan;
}());
exports.Scan = Scan;
