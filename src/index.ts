#!/usr/bin/env node

import { Config } from './Config';
import { Init } from './Init';
import { Migrate } from './Migrate';
import { Scan } from './Scan';
import program from 'commander';

// Program
program.version('0.0.1');

// Commands
program.command('init').action(command => Init.run());
program.command('scan').action(command => Scan.run(Config.get()));
program.command('migrate').action(command => Migrate.run(Config.get()));

// Parse command
program.parse(process.argv);

// Help
if (!program.args.length) program.help();
