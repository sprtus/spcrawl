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
var sp_taxonomy_1 = require("@pnp/sp-taxonomy");
var pnpjs_1 = require("@pnp/pnpjs");
var chalk_1 = __importDefault(require("chalk"));
// Migrate data
var departments = [
    { name: 'Administration', id: 25 },
    { name: 'Sales, Marketing ＆ New Product Innovation', id: 14 },
    { name: 'Architecture ＆ Planning', id: 27 },
    { name: 'Assistant General Counsel', id: 19 },
    { name: 'C＆EN', id: 13 },
    { name: 'Membership ＆ Society Services', id: 6 },
    { name: 'Treasurer ＆ CFO', id: 22 },
    { name: 'Sales, Marketing ＆ New Product Innovation', id: 16 },
    { name: 'Publications', id: 45 },
    { name: 'Education', id: 32 },
    { name: 'Employee Services', id: 10 },
    { name: 'Education', id: 1 },
    { name: 'ACS Executive Leadership Team', id: 31 },
    { name: 'Finance', id: 23 },
    { name: 'Marketing ＆ Sales', id: 5 },
    { name: 'Green Chemistry Institute', id: 9 },
    { name: 'Education', id: 2 },
    { name: 'HRIS', id: 11 },
    { name: 'Human Resources', id: 35 },
    { name: 'Secretary ＆ General Counsel', id: 18 },
    { name: 'Meetings ＆ Exposition Services', id: 8 },
    { name: 'Investments', id: 21 },
    { name: 'Journals Publishing Group', id: 12 },
    { name: 'Education', id: 3 },
    { name: 'Meetings ＆ Exposition Services', id: 34 },
    { name: 'Member Programs ＆ Communities', id: 7 },
    { name: 'Membership ＆ Society Services', id: 33 },
    { name: 'External Affairs ＆ Communications', id: 20 },
    { name: 'Organization Development', id: 43 },
    { name: 'Learning and Career Development', id: 42 },
    { name: 'Learning and Career Development', id: 4 },
    { name: 'Publications Business Support', id: 26 },
    { name: 'Publications', id: 36 },
    { name: 'Research Grants', id: 46 },
    { name: 'Digital Strategy ＆ Publishing Platforms', id: 15 },
    { name: 'Scientific Advancement', id: 17 },
    { name: 'Secretary ＆ General Counsel', id: 37 },
    { name: 'Secretary ＆ General Counsel', id: 41 },
    { name: 'Society ＆ Administrative Technology', id: 29 },
    { name: 'Software ＆ Engineering', id: 28 },
    { name: 'Technical Infrastructure Services', id: 48 },
    { name: 'Technical Programs ＆ Activities', id: 47 },
    { name: 'Total Rewards', id: 44 },
    { name: 'Treasurer ＆ CFO', id: 38 },
    { name: 'Publications', id: 49 },
    { name: 'Washington IT Operations', id: 39 },
    { name: 'Web Strategy ＆ Operations', id: 30 },
];
var divisions = [
    { name: 'CAS', id: 1 },
    { name: 'Education', id: 2 },
    { name: 'ACS Executive Leadership Team', id: 10 },
    { name: 'Human Resources', id: 4 },
    { name: 'Membership ＆ Society Services', id: 3 },
    { name: 'Publications', id: 5 },
    { name: 'Scientific Advancement', id: 11 },
    { name: 'Secretary ＆ General Counsel', id: 7 },
    { name: 'Treasurer ＆ CFO', id: 8 },
    { name: 'Washington IT Operations', id: 9 },
];
var formCategories = [
    { name: 'Administration', id: 1 },
    { name: 'Finance', id: 2 },
    { name: 'Human Resources', id: 3 },
    { name: 'Washington IT Operations', id: 4 },
    { name: 'Meetings', id: 6 },
    { name: 'Membership', id: 5 },
    { name: 'Publications', id: 7 },
    { name: 'Travel', id: 8 },
];
var formTypes = [
    { name: 'Form', id: 1 },
    { name: 'Instructions', id: 2 },
    { name: 'Policy', id: 3 },
];
var offices = [
    { name: 'Accounts Payable', id: 'Accounts Payable' },
    { name: 'Benefits', id: 'Benefits' },
    { name: 'Budget ＆ Analysis', id: 'Budgets & Analysis' },
    { name: 'Conferencing ＆ Webinars', id: 'Conferencing' },
    { name: 'Contracts', id: 'Contract Administration' },
    { name: 'Finance', id: 'Finance' },
    { name: 'Financial Services', id: 'General Accounting' },
    { name: 'Mail, Print, ＆ Copy Services', id: 'Copy Center' },
    { name: 'National Meetings', id: 'National Meetings' },
    { name: 'Payroll', id: 'Payroll' },
    { name: 'Purchasing', id: 'Purchasing' },
    { name: 'Taxes', id: 'Taxes' },
];
// "Tags" term set ID
var termSetId = 'a664dbf2-f7d4-485e-820a-22757e1418e9';
// Migrate ACS
var Tag = /** @class */ (function () {
    function Tag() {
    }
    Tag.run = function (config) {
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
                                                var store, termset, terms, _loop_2, i, _loop_3, i, _loop_4, i, _loop_5, i, _loop_6, i, web, tagsField, documents, _i, _a, document_1, locationTag, tags, departmentId, _loop_7, _b, departments_1, department, state_1, divisionId, _loop_8, _c, divisions_1, division, state_2, formCategoryId, _loop_9, _d, formCategories_1, formCategory, state_3, formTypeId, _loop_10, _e, formTypes_1, formType, state_4, _loop_11, _f, offices_1, office, state_5, termString, _g, tags_1, tag, updateData;
                                                return __generator(this, function (_h) {
                                                    switch (_h.label) {
                                                        case 0:
                                                            console.log("  Connected to " + settings.siteUrl);
                                                            // Get term set
                                                            console.log('Getting term store...');
                                                            return [4 /*yield*/, sp_taxonomy_1.taxonomy.getDefaultSiteCollectionTermStore().get()];
                                                        case 1:
                                                            store = _h.sent();
                                                            return [4 /*yield*/, store.getTermSetById(termSetId).get()];
                                                        case 2:
                                                            termset = _h.sent();
                                                            console.log("  " + store.Name + "/" + termSetId);
                                                            return [4 /*yield*/, termset.terms.get()];
                                                        case 3:
                                                            terms = _h.sent();
                                                            console.log("  " + terms.length + " terms\n");
                                                            _loop_2 = function (i) {
                                                                // Find matching terms
                                                                var matchingTerms = terms.filter(function (term) { return term.Name == departments[i].name; });
                                                                // Not found
                                                                if (!matchingTerms.length) {
                                                                    console.log(chalk_1.default.bold(departments[i].name + ":") + " " + chalk_1.default.red('not found'));
                                                                    return "continue";
                                                                }
                                                                // Update term
                                                                var guid = matchingTerms[0].Id;
                                                                guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
                                                                departments[i].guid = guid;
                                                                departments[i].term = matchingTerms[0];
                                                                console.log(chalk_1.default.bold(departments[i].name + ":") + " " + chalk_1.default.green(departments[i].guid));
                                                            };
                                                            // Departments
                                                            for (i = 0; i < departments.length; i++) {
                                                                _loop_2(i);
                                                            }
                                                            console.log('');
                                                            _loop_3 = function (i) {
                                                                // Find matching terms
                                                                var matchingTerms = terms.filter(function (term) { return term.Name == divisions[i].name; });
                                                                // Not found
                                                                if (!matchingTerms.length) {
                                                                    console.log(chalk_1.default.bold(divisions[i].name + ":") + " " + chalk_1.default.red('not found'));
                                                                    return "continue";
                                                                }
                                                                // Update term
                                                                var guid = matchingTerms[0].Id;
                                                                guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
                                                                divisions[i].guid = guid;
                                                                divisions[i].term = matchingTerms[0];
                                                                console.log(chalk_1.default.bold(divisions[i].name + ":") + " " + chalk_1.default.green(divisions[i].guid));
                                                            };
                                                            // Divisions
                                                            for (i = 0; i < divisions.length; i++) {
                                                                _loop_3(i);
                                                            }
                                                            console.log('');
                                                            _loop_4 = function (i) {
                                                                // Find matching terms
                                                                var matchingTerms = terms.filter(function (term) { return term.Name == formCategories[i].name; });
                                                                // Not found
                                                                if (!matchingTerms.length) {
                                                                    console.log(chalk_1.default.bold(formCategories[i].name + ":") + " " + chalk_1.default.red('not found'));
                                                                    return "continue";
                                                                }
                                                                // Update term
                                                                var guid = matchingTerms[0].Id;
                                                                guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
                                                                formCategories[i].guid = guid;
                                                                formCategories[i].term = matchingTerms[0];
                                                                console.log(chalk_1.default.bold(formCategories[i].name + ":") + " " + chalk_1.default.green(formCategories[i].guid));
                                                            };
                                                            // Form categories
                                                            for (i = 0; i < formCategories.length; i++) {
                                                                _loop_4(i);
                                                            }
                                                            console.log('');
                                                            _loop_5 = function (i) {
                                                                // Find matching terms
                                                                var matchingTerms = terms.filter(function (term) { return term.Name == formTypes[i].name; });
                                                                // Not found
                                                                if (!matchingTerms.length) {
                                                                    console.log(chalk_1.default.bold(formTypes[i].name + ":") + " " + chalk_1.default.red('not found'));
                                                                    return "continue";
                                                                }
                                                                // Update term
                                                                var guid = matchingTerms[0].Id;
                                                                guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
                                                                formTypes[i].guid = guid;
                                                                formTypes[i].term = matchingTerms[0];
                                                                console.log(chalk_1.default.bold(formTypes[i].name + ":") + " " + chalk_1.default.green(formTypes[i].guid));
                                                            };
                                                            // Form categories
                                                            for (i = 0; i < formTypes.length; i++) {
                                                                _loop_5(i);
                                                            }
                                                            console.log('');
                                                            _loop_6 = function (i) {
                                                                // Find matching terms
                                                                var matchingTerms = terms.filter(function (term) { return term.Name == offices[i].name; });
                                                                // Not found
                                                                if (!matchingTerms.length) {
                                                                    console.log(chalk_1.default.bold(offices[i].name + ":") + " " + chalk_1.default.red('not found'));
                                                                    return "continue";
                                                                }
                                                                // Update term
                                                                var guid = matchingTerms[0].Id;
                                                                guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
                                                                offices[i].guid = guid;
                                                                offices[i].term = matchingTerms[0];
                                                                console.log(chalk_1.default.bold(offices[i].name + ":") + " " + chalk_1.default.green(offices[i].guid));
                                                            };
                                                            // Offices
                                                            for (i = 0; i < offices.length; i++) {
                                                                _loop_6(i);
                                                            }
                                                            console.log('');
                                                            web = new pnpjs_1.Web(siteUrl + "/policies-forms");
                                                            return [4 /*yield*/, web.lists.getByTitle('Documents').fields.getByTitle('ACS_Tags_0').get()];
                                                        case 4:
                                                            tagsField = _h.sent();
                                                            return [4 /*yield*/, web.lists.getByTitle('Documents').items.select('ID', 'Title', 'Migrate_Department', 'Migrate_Division', 'Migrate_FormCategory', 'Migrate_FormType', 'Migrate_Office', 'Keywords', 'File/Name', 'File/ServerRelativeUrl', 'File/Title', 'File/Length').expand('file').getAll().catch(function (e) { return console.error; })];
                                                        case 5:
                                                            documents = _h.sent();
                                                            _i = 0, _a = documents;
                                                            _h.label = 6;
                                                        case 6:
                                                            if (!(_i < _a.length)) return [3 /*break*/, 10];
                                                            document_1 = _a[_i];
                                                            if (!document_1.File || !document_1.File.ServerRelativeUrl)
                                                                return [3 /*break*/, 9];
                                                            console.log("Updating " + document_1.File.ServerRelativeUrl);
                                                            locationTag = divisions[2];
                                                            tags = [];
                                                            // Department tags
                                                            if (document_1.Migrate_Department) {
                                                                departmentId = parseInt(document_1.Migrate_Department);
                                                                _loop_7 = function (department) {
                                                                    if (department.id == departmentId && !tags.filter(function (tag) { return tag.guid == department.guid; }).length) {
                                                                        tags.push(department);
                                                                        return "break";
                                                                    }
                                                                };
                                                                for (_b = 0, departments_1 = departments; _b < departments_1.length; _b++) {
                                                                    department = departments_1[_b];
                                                                    state_1 = _loop_7(department);
                                                                    if (state_1 === "break")
                                                                        break;
                                                                }
                                                            }
                                                            // Division tags
                                                            if (document_1.Migrate_Division) {
                                                                divisionId = parseInt(document_1.Migrate_Division);
                                                                _loop_8 = function (division) {
                                                                    if (division.id == divisionId && !tags.filter(function (tag) { return tag.guid == division.guid; }).length) {
                                                                        tags.push(division);
                                                                        return "break";
                                                                    }
                                                                };
                                                                for (_c = 0, divisions_1 = divisions; _c < divisions_1.length; _c++) {
                                                                    division = divisions_1[_c];
                                                                    state_2 = _loop_8(division);
                                                                    if (state_2 === "break")
                                                                        break;
                                                                }
                                                            }
                                                            // Form category tags
                                                            if (document_1.Migrate_FormCategory) {
                                                                formCategoryId = parseInt(document_1.Migrate_FormCategory);
                                                                _loop_9 = function (formCategory) {
                                                                    if (formCategory.id == formCategoryId && !tags.filter(function (tag) { return tag.guid == formCategory.guid; }).length) {
                                                                        tags.push(formCategory);
                                                                        return "break";
                                                                    }
                                                                };
                                                                for (_d = 0, formCategories_1 = formCategories; _d < formCategories_1.length; _d++) {
                                                                    formCategory = formCategories_1[_d];
                                                                    state_3 = _loop_9(formCategory);
                                                                    if (state_3 === "break")
                                                                        break;
                                                                }
                                                            }
                                                            // Form type tags
                                                            if (document_1.Migrate_FormType) {
                                                                formTypeId = parseInt(document_1.Migrate_FormType);
                                                                _loop_10 = function (formType) {
                                                                    if (formType.id == formTypeId && !tags.filter(function (tag) { return tag.guid == formType.guid; }).length) {
                                                                        tags.push(formType);
                                                                        return "break";
                                                                    }
                                                                };
                                                                for (_e = 0, formTypes_1 = formTypes; _e < formTypes_1.length; _e++) {
                                                                    formType = formTypes_1[_e];
                                                                    state_4 = _loop_10(formType);
                                                                    if (state_4 === "break")
                                                                        break;
                                                                }
                                                            }
                                                            // Office tags
                                                            if (document_1.Migrate_Office) {
                                                                _loop_11 = function (office) {
                                                                    if (office.id == document_1.Migrate_Office && !tags.filter(function (tag) { return tag.guid == office.guid; }).length) {
                                                                        tags.push(office);
                                                                        return "break";
                                                                    }
                                                                };
                                                                for (_f = 0, offices_1 = offices; _f < offices_1.length; _f++) {
                                                                    office = offices_1[_f];
                                                                    state_5 = _loop_11(office);
                                                                    if (state_5 === "break")
                                                                        break;
                                                                }
                                                            }
                                                            if (!tags.length) return [3 /*break*/, 8];
                                                            console.log("  " + chalk_1.default.cyan(tags.map(function (tag) { return tag.name; }).join(', ')));
                                                            termString = '';
                                                            for (_g = 0, tags_1 = tags; _g < tags_1.length; _g++) {
                                                                tag = tags_1[_g];
                                                                termString += "-1;#" + tag.term.Name + "|" + tag.guid + ";#";
                                                            }
                                                            updateData = {};
                                                            updateData[tagsField.InternalName] = termString;
                                                            // Check out
                                                            // await web.getFileByServerRelativeUrl(document.File.ServerRelativeUrl).checkout().catch(e => {});
                                                            // console.log(chalk.gray('  Checked out'));
                                                            // Update tags
                                                            return [4 /*yield*/, web.lists.getByTitle('Documents').items.getById(document_1.ID).update(updateData).catch(function (e) { return console.error; })];
                                                        case 7:
                                                            // Check out
                                                            // await web.getFileByServerRelativeUrl(document.File.ServerRelativeUrl).checkout().catch(e => {});
                                                            // console.log(chalk.gray('  Checked out'));
                                                            // Update tags
                                                            _h.sent();
                                                            console.log(chalk_1.default.gray('  Updated'));
                                                            return [3 /*break*/, 9];
                                                        case 8:
                                                            console.log(chalk_1.default.gray('  No tags'));
                                                            _h.label = 9;
                                                        case 9:
                                                            _i++;
                                                            return [3 /*break*/, 6];
                                                        case 10: return [2 /*return*/];
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
    return Tag;
}());
exports.Tag = Tag;
