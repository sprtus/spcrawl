import { ICrawlOptions } from './Config';
import { Utility } from './Utility';
import chalk from 'chalk';
import fs from 'fs';

// Configuration file template
const configTemplate: ICrawlOptions = {
  auth: null,
  sites:[
    'https://yoursite.sharepoint.com',
  ],
};

export class Init {
  static run(): void {
    // Begin
    console.log(chalk.cyan('Initializing SPCrawl...'));

    // Get config file path
    const configFilePath: string = Utility.path('spcrawl.json');
    console.log(`  - Creating config file: ${chalk.whiteBright(configFilePath)}`)

    // File exists
    if (fs.existsSync(configFilePath)) {
      console.log(chalk.red(`  - ${chalk.redBright.underline('spcrawl.json')} already exists`));
      throw new Error();
    }

    // Create file
    try {
      fs.writeFileSync(configFilePath, JSON.stringify(configTemplate, null, 2), { flag: 'wx' });
    } catch (e) {
      console.error(e);
    }

    // Done
    console.log(chalk.greenBright('Done.'));
  }
};
