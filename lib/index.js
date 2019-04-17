#!/usr/bin/env node
"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var Config_1 = require("./Config");
var Init_1 = require("./Init");
var News_1 = require("./News");
var Phoenix_1 = require("./Phoenix");
var Policies_1 = require("./Policies");
var Scan_1 = require("./Scan");
var Tag_1 = require("./Tag");
var commander_1 = __importDefault(require("commander"));
// Program
commander_1.default.version('0.0.1');
// Commands
commander_1.default.command('init').action(function (command) { return Init_1.Init.run(); });
commander_1.default.command('news-crawl').action(function (command) { return News_1.News.crawl(Config_1.Config.get()); });
commander_1.default.command('news-create').action(function (command) { return News_1.News.create(Config_1.Config.get()); });
commander_1.default.command('phoenix-crawl').action(function (command) { return Phoenix_1.Phoenix.crawl(Config_1.Config.get()); });
commander_1.default.command('phoenix-create').action(function (command) { return Phoenix_1.Phoenix.create(Config_1.Config.get()); });
commander_1.default.command('policies').action(function (command) { return Policies_1.Policies.run(Config_1.Config.get()); });
commander_1.default.command('scan').action(function (command) { return Scan_1.Scan.run(Config_1.Config.get()); });
commander_1.default.command('tag').action(function (command) { return Tag_1.Tag.run(Config_1.Config.get()); });
// Parse command
commander_1.default.parse(process.argv);
// Help
if (!commander_1.default.args.length)
    commander_1.default.help();
