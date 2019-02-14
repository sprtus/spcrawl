"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var Utility_1 = require("./Utility");
var chalk_1 = __importDefault(require("chalk"));
var fs_1 = __importDefault(require("fs"));
// Configuration file template
var configTemplate = {
    auth: null,
    sites: [
        'https://yoursite.sharepoint.com',
    ],
};
var Init = /** @class */ (function () {
    function Init() {
    }
    Init.run = function () {
        // Begin
        console.log(chalk_1.default.cyan('Initializing SPCrawl...'));
        // Get config file path
        var configFilePath = Utility_1.Utility.path('spcrawl.json');
        console.log("  - Creating config file: " + chalk_1.default.whiteBright(configFilePath));
        // File exists
        if (fs_1.default.existsSync(configFilePath)) {
            console.log(chalk_1.default.red("  - " + chalk_1.default.redBright.underline('spcrawl.json') + " already exists"));
            throw new Error();
        }
        // Create file
        try {
            fs_1.default.writeFileSync(configFilePath, JSON.stringify(configTemplate, null, 2), { flag: 'wx' });
        }
        catch (e) {
            console.error(e);
        }
        // Done
        console.log(chalk_1.default.greenBright('Done.'));
    };
    return Init;
}());
exports.Init = Init;
;
