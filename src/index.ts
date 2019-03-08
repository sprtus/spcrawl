#!/usr/bin/env node

import { Config } from './Config';
import { Init } from './Init';
import { News } from './News';
import { Policies } from './Policies';
import { Scan } from './Scan';
import { Tag } from './Tag';
import program from 'commander';

// Program
program.version('0.0.1');

// Commands
program.command('init').action(command => Init.run());
program.command('news').action(command => News.run(Config.get()));
program.command('policies').action(command => Policies.run(Config.get()));
program.command('scan').action(command => Scan.run(Config.get()));
program.command('tag').action(command => Tag.run(Config.get()));

// Parse command
program.parse(process.argv);

// Help
if (!program.args.length) program.help();
