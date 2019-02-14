import { IAuthOptions } from 'node-sp-auth';
import { Utility } from './Utility';
import chalk from 'chalk';
import fs from 'fs';

export class Config {
  // Get configuration settings
  static get(configFilePath: string = Utility.path('spcrawl.json')): ICrawlOptions {
    // No config file
    if (!fs.existsSync(configFilePath)) {
      console.log(chalk.red(`Config file not found: ${configFilePath}`))
      console.log(`Use ${chalk.yellowBright.underline('crawlsp init')} to create a new config file.`);
      throw new Error();
    }

    // Get config JSON
    let config: ICrawlOptions;
    try {
      const fileContents: any = fs.readFileSync(configFilePath);
      config = JSON.parse(fileContents);
    } catch (e) {
      throw new Error(e);
    }

    return config;
  }
};

export interface ICrawlOptions {
  auth: IAuthOptions | null;
  sites: string[];
}
