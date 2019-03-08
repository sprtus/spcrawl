import _ from 'lodash';
import { ICrawlOptions } from './Config';
import { PnpNode } from 'sp-pnp-node';
import { Web, CheckinType } from '@pnp/pnpjs';
import chalk from 'chalk';

// Migrate ACS news
export class News {
  static async run (config: ICrawlOptions): Promise<void> {
    // Crawl sites
    for (const siteUrl of config.sites) {
      console.log(`Connecting to ${siteUrl}`);

      // Connect
      await new PnpNode().init().then(async settings => {
        console.log(`  Connected to ${settings.siteUrl}`);

        // Get web
        const web: Web = new Web(`${siteUrl}/Publications/SalesMarketing/WebInnovation`);

        // Get all documents
        console.log(`Getting documents: ${siteUrl}/Publications/SalesMarketing/WebInnovation/StatusReports`)
        const documents = await web.lists.getByTitle('StatusReports').items.select('ID', 'Title', 'Keywords', 'File/Name', 'File/ServerRelativeUrl', 'File/Title').expand('file').getAll().catch(e => console.error);

        // Iterate pages
        for (const document of documents as any[]) {
          if (!document.File || !document.File.ServerRelativeUrl) continue;
          console.log(`Updating ${document.File.ServerRelativeUrl}`);

        }
      });
    }

    console.log('Done.');
  }
}
