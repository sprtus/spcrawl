#!/usr/bin/env node
"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var Config_1 = require("./Config");
var Init_1 = require("./Init");
var Migrate_1 = require("./Migrate");
var Scan_1 = require("./Scan");
var commander_1 = __importDefault(require("commander"));
// Program
commander_1.default.version('0.0.1');
// Commands
commander_1.default.command('init').action(function (command) { return Init_1.Init.run(); });
commander_1.default.command('scan').action(function (command) { return Scan_1.Scan.run(Config_1.Config.get()); });
commander_1.default.command('migrate').action(function (command) { return Migrate_1.Migrate.run(Config_1.Config.get()); });
// Parse command
commander_1.default.parse(process.argv);
// Help
if (!commander_1.default.args.length)
    commander_1.default.help();
