"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var Utility_1 = require("./Utility");
var chalk_1 = __importDefault(require("chalk"));
var fs_1 = __importDefault(require("fs"));
var Config = /** @class */ (function () {
    function Config() {
    }
    // Get configuration settings
    Config.get = function (configFilePath) {
        if (configFilePath === void 0) { configFilePath = Utility_1.Utility.path('spcrawl.json'); }
        // No config file
        if (!fs_1.default.existsSync(configFilePath)) {
            console.log(chalk_1.default.red("Config file not found: " + configFilePath));
            console.log("Use " + chalk_1.default.yellowBright.underline('crawlsp init') + " to create a new config file.");
            throw new Error();
        }
        // Get config JSON
        var config;
        try {
            var fileContents = fs_1.default.readFileSync(configFilePath);
            config = JSON.parse(fileContents);
        }
        catch (e) {
            throw new Error(e);
        }
        return config;
    };
    return Config;
}());
exports.Config = Config;
;
