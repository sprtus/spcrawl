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
var sp_pnp_node_1 = require("sp-pnp-node");
var Utility_1 = require("./Utility");
var pnpjs_1 = require("@pnp/pnpjs");
var chalk_1 = __importDefault(require("chalk"));
var fs_1 = __importDefault(require("fs"));
// Migrate ACS news
var Phoenix = /** @class */ (function () {
    function Phoenix() {
    }
    Phoenix.crawl = function (config) {
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
                                                var articles, web, files, _i, _a, file, fileName, d, duplicate, title, pageName;
                                                return __generator(this, function (_b) {
                                                    switch (_b.label) {
                                                        case 0:
                                                            console.log("  Connected to " + settings.siteUrl);
                                                            articles = [];
                                                            web = new pnpjs_1.Web(siteUrl + "/phoenix");
                                                            // Get all documents
                                                            console.log("\nGetting documents");
                                                            return [4 /*yield*/, web.getFolderByServerRelativeUrl("/phoenix/The Phoenix Archives").files.select('ServerRelativeUrl', 'Title', 'Name', 'TimeCreated', 'Month', 'Year').top(5000).get().catch(function (e) { return console.error; })];
                                                        case 1:
                                                            files = _b.sent();
                                                            // Iterate pages
                                                            for (_i = 0, _a = files; _i < _a.length; _i++) {
                                                                file = _a[_i];
                                                                if (!file.ServerRelativeUrl)
                                                                    continue;
                                                                console.log("  " + file.Name);
                                                                fileName = file.Name.split('.')[0];
                                                                d = {
                                                                    Year: null,
                                                                    Month: null,
                                                                    MonthText: null,
                                                                };
                                                                duplicate = false;
                                                                if (fileName.split('_').length === 3) {
                                                                    d.Year = parseInt(fileName.split('_')[0]);
                                                                    d.Month = parseInt(fileName.split('_')[1]);
                                                                    d.MonthText = fileName.split('_')[2];
                                                                    if (d.MonthText.indexOf('-') !== -1) {
                                                                        d.MonthText = d.MonthText.split('-')[0];
                                                                        duplicate = true;
                                                                    }
                                                                }
                                                                title = "The Phoenix: " + (d.Year ? d.MonthText + ", " + d.Year : fileName) + (duplicate ? ' (2)' : '');
                                                                pageName = "The-Phoenix-" + (d.Year ? d.Year + "-" + d.Month + (duplicate ? '-2' : '') : fileName) + ".aspx";
                                                                // Save article
                                                                articles.push({
                                                                    Title: title,
                                                                    PageName: pageName,
                                                                    Date: d,
                                                                    Duplicate: duplicate,
                                                                    Url: "" + siteUrl + file.ServerRelativeUrl,
                                                                    File: file,
                                                                });
                                                            }
                                                            // Write to file
                                                            console.log("\nSaving " + articles.length + " articles...\n  " + chalk_1.default.greenBright(Utility_1.Utility.path('phoenix.json')));
                                                            try {
                                                                fs_1.default.writeFileSync(Utility_1.Utility.path('phoenix.json'), JSON.stringify(articles, null, 2), { flag: 'w' });
                                                            }
                                                            catch (e) {
                                                                console.error(e);
                                                            }
                                                            return [2 /*return*/];
                                                    }
                                                });
                                            }); })];
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
    Phoenix.create = function (config) {
        return __awaiter(this, void 0, void 0, function () {
            var newsSiteUrl;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        newsSiteUrl = 'https://nucleusdev.acs.org/news-events';
                        console.log("Connecting to " + newsSiteUrl);
                        // Connect
                        return [4 /*yield*/, new sp_pnp_node_1.PnpNode({
                                siteUrl: newsSiteUrl,
                            }).init().then(function (settings) { return __awaiter(_this, void 0, void 0, function () {
                                var articles, web, _i, articles_1, article, relativePageUrl, pageContent, checkedOut, page, fileCheckedIn, fileApproved;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            console.log("  Connected to " + settings.siteUrl + "\n");
                                            articles = fs_1.default.readFileSync(Utility_1.Utility.path('./phoenix.json'));
                                            articles = JSON.parse(articles);
                                            web = new pnpjs_1.Web("" + newsSiteUrl);
                                            _i = 0, articles_1 = articles;
                                            _a.label = 1;
                                        case 1:
                                            if (!(_i < articles_1.length)) return [3 /*break*/, 7];
                                            article = articles_1[_i];
                                            relativePageUrl = "/news-events/Pages/" + article.PageName;
                                            console.log("Creating: " + chalk_1.default.cyan(relativePageUrl));
                                            pageContent = "<%@ Page Inherits=\"Microsoft.SharePoint.Publishing.TemplateRedirectionPage,Microsoft.SharePoint.Publishing,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c\" %> <%@ Reference VirtualPath=\"~TemplatePageUrl\" %> <%@ Reference VirtualPath=\"~masterurl/custom.master\" %><%@ Register Tagprefix=\"SharePoint\" Namespace=\"Microsoft.SharePoint.WebControls\" Assembly=\"Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" %>\n        <html xmlns:mso=\"urn:schemas-microsoft-com:office:office\" xmlns:msdt=\"uuid:C2F41010-65B3-11d1-A29F-00AA00C14882\"><head>\n        <!--[if gte mso 9]><SharePoint:CTFieldRefs runat=server Prefix=\"mso:\" FieldList=\"FileLeafRef,Comments,PublishingStartDate,PublishingExpirationDate,PublishingContactEmail,PublishingContactName,PublishingContactPicture,PublishingPageLayout,PublishingVariationGroupID,PublishingVariationRelationshipLinkFieldID,PublishingRollupImage,Audience,PublishingIsFurlPage,SeoBrowserTitle,SeoMetaDescription,SeoKeywords,RobotsNoIndex,ACS_Summary,PublishingPageContent,ACS_Featured,odea017de2994eadb29fca35aaab7db7,TaxCatchAllLabel,ACS_Thumbnail,ACS_ShowSideNav,ACS_DivisionLanding,ACS_TeamLanding,ACS_EntryDate,ACS_Byline,n53db4776dca48dba0972095b72739fe,ACS_DigestSend,ACS_DigestOrder,ACS_DigestPubStart,ACS_DigestPubEnd,ACS_WhoToCallListName\"><xml>\n        <mso:CustomDocumentProperties>\n        <mso:PublishingContact msdt:dt=\"string\">778</mso:PublishingContact>\n        <mso:PublishingIsFurlPage msdt:dt=\"string\">0</mso:PublishingIsFurlPage>\n        <mso:display_urn_x003a_schemas-microsoft-com_x003a_office_x003a_office_x0023_PublishingContact msdt:dt=\"string\">WFService, SharePoint</mso:display_urn_x003a_schemas-microsoft-com_x003a_office_x003a_office_x0023_PublishingContact>\n        <mso:ACS_DigestSend msdt:dt=\"string\">1</mso:ACS_DigestSend>\n        <mso:PublishingContactPicture msdt:dt=\"string\"></mso:PublishingContactPicture>\n        <mso:RobotsNoIndex msdt:dt=\"string\">0</mso:RobotsNoIndex>\n        <mso:ACS_ShowSideNav msdt:dt=\"string\">0</mso:ACS_ShowSideNav>\n        <mso:ACS_DivisionLanding msdt:dt=\"string\">0</mso:ACS_DivisionLanding>\n        <mso:ACS_EntryDate msdt:dt=\"string\">" + (article.Date.Year ? article.Date.Year + "-" + article.Date.Month + "-01T00:00:00Z" : article.File.TimeCreated) + "</mso:ACS_EntryDate>\n        <mso:ACS_Send msdt:dt=\"string\">0</mso:ACS_Send>\n        <mso:ContentTypeId msdt:dt=\"string\">0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF390059178973DC71AF439E03FE69E2CBC20F02002748B97C4D70AD46A057697FC45223CB</mso:ContentTypeId>\n        <mso:PublishingContactName msdt:dt=\"string\"></mso:PublishingContactName>\n        <mso:ACS_Featured msdt:dt=\"string\">0</mso:ACS_Featured>\n        <mso:PublishingPageLayoutName msdt:dt=\"string\">News.aspx</mso:PublishingPageLayoutName>\n        <mso:Comments msdt:dt=\"string\"></mso:Comments>\n        <mso:PublishingContactEmail msdt:dt=\"string\"></mso:PublishingContactEmail>\n        <mso:PublishingPageLayout msdt:dt=\"string\">https://nucleusdev.acs.org/_catalogs/masterpage/acs-nucleus/News.aspx, News</mso:PublishingPageLayout>\n        <mso:ACS_TeamLanding msdt:dt=\"string\">0</mso:ACS_TeamLanding>\n        <mso:PublishingPageContent msdt:dt=\"string\">&lt;p&gt;&lt;strong&gt;The Phoenix&lt;/strong&gt;" + (article.Date.Year ? " issue for " + article.Date.MonthText + ", " + article.Date.Year.toString() : '') + " is available. Click on the link below to view the newsletter.&lt;/p&gt;&lt;p class=&quot;ms-rteElement-ButtonRowPrimary&quot&gt;&lt;a href=&quot;" + article.Url + "&quot;&gt;[icon:far fa-file-pdf] " + article.File.Name + "&lt;/a&gt;&lt;/p&gt;</mso:PublishingPageContent>\n        <mso:ACS_DigestPubStart msdt:dt=\"string\"></mso:ACS_DigestPubStart>\n        <mso:ACS_DigestOrder msdt:dt=\"string\"></mso:ACS_DigestOrder>\n        <mso:ACS_Summary msdt:dt=\"string\">The" + (article.Date.Year ? " " + article.Date.MonthText + ", " + article.Date.Year.toString() + " edition of" : '') + " The Phoenix</mso:ACS_Summary>\n        <mso:ACS_DigestPubEnd msdt:dt=\"string\"></mso:ACS_DigestPubEnd>\n        <mso:Order msdt:dt=\"string\"></mso:Order>\n        <mso:ACS_Tags msdt:dt=\"string\"></mso:ACS_Tags>\n        <mso:odea017de2994eadb29fca35aaab7db7 msdt:dt=\"string\"></mso:odea017de2994eadb29fca35aaab7db7>\n        <mso:n53db4776dca48dba0972095b72739fe msdt:dt=\"string\">The Phoenix|36b4dd5f-3571-44be-a225-6154c18f480f</mso:n53db4776dca48dba0972095b72739fe>\n        <mso:ACS_NewsChannel msdt:dt=\"string\">58;#The Phoenix|36b4dd5f-3571-44be-a225-6154c18f480f</mso:ACS_NewsChannel>\n        <mso:TaxCatchAll msdt:dt=\"string\">58;#The Phoenix|36b4dd5f-3571-44be-a225-6154c18f480f</mso:TaxCatchAll>\n        </mso:CustomDocumentProperties>\n        </xml></SharePoint:CTFieldRefs><![endif]-->\n        <title>" + article.Title + "</title></head>";
                                            return [4 /*yield*/, web.getFileByServerRelativeUrl(relativePageUrl).checkout().catch(function (e) { })];
                                        case 2:
                                            checkedOut = _a.sent();
                                            console.log(chalk_1.default.gray('  Checked out'));
                                            return [4 /*yield*/, web.getFolderByServerRelativeUrl('/news-events/Pages').files.add(article.PageName, pageContent, true).catch(function (e) { return console.error; })];
                                        case 3:
                                            page = _a.sent();
                                            console.log(chalk_1.default.gray('  Updated'));
                                            return [4 /*yield*/, web.getFileByServerRelativeUrl(relativePageUrl).checkin('Checked in via automated migration', pnpjs_1.CheckinType.Major).catch(function (e) { })];
                                        case 4:
                                            fileCheckedIn = _a.sent();
                                            console.log(chalk_1.default.gray('  Checked in'));
                                            return [4 /*yield*/, web.getFileByServerRelativeUrl(relativePageUrl).approve('Approved via automated migration').catch(function (e) { })];
                                        case 5:
                                            fileApproved = _a.sent();
                                            console.log(chalk_1.default.gray('  Approved'));
                                            _a.label = 6;
                                        case 6:
                                            _i++;
                                            return [3 /*break*/, 1];
                                        case 7: return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 1:
                        // Connect
                        _a.sent();
                        console.log('\nDone.');
                        return [2 /*return*/];
                }
            });
        });
    };
    return Phoenix;
}());
exports.Phoenix = Phoenix;
