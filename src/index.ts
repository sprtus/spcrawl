#!/usr/bin/env node

import { Config } from './Config';
import { Init } from './Init';
import { News } from './News';
import { Phoenix } from './Phoenix';
import { Policies } from './Policies';
import { Scan } from './Scan';
import { Tag } from './Tag';
import program from 'commander';

// Program
program.version('0.0.1');

// Commands
program.command('init').action(command => Init.run());
program.command('news-crawl').action(command => News.crawl(Config.get()));
program.command('news-create').action(command => News.create(Config.get()));
program.command('phoenix-crawl').action(command => Phoenix.crawl(Config.get()));
program.command('phoenix-create').action(command => Phoenix.create(Config.get()));
program.command('policies').action(command => Policies.run(Config.get()));
program.command('scan').action(command => Scan.run(Config.get()));
program.command('tag').action(command => Tag.run(Config.get()));

// Parse command
program.parse(process.argv);

// Help
if (!program.args.length) program.help();
