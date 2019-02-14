"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var path_1 = __importDefault(require("path"));
var Utility = /** @class */ (function () {
    function Utility() {
    }
    // Build path
    Utility.path = function (filePath, cwd) {
        if (cwd === void 0) { cwd = true; }
        return path_1.default.normalize((cwd ? process.cwd() : __dirname) + "/" + filePath);
    };
    return Utility;
}());
exports.Utility = Utility;
;
