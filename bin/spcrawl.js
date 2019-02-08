#!/usr/bin/env node

const program = require('commander');
const config = require('../src/config');

// Commands
const init = require('../src/init');
const crawl = require('../src/crawl');

// Version
program.version('1.0.0');

// Init
program.command('init').action(cmd => init());

// Crawl
program.command('crawl').action(cmd => crawl(config()));

// Parse command
program.parse(process.argv);

// Help
if (!program.args.length) program.help();
