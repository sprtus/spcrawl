import _ from 'lodash';
import { ICrawlOptions } from './Config';
import { ISite } from './data/ISite';
import { NewContentType, IContentType } from './data/IContentType';
import { NewField, IField } from './data/IField';
import { NewWeb } from './data/IWeb';
import { PnpNode } from 'sp-pnp-node';
import { Utility } from './Utility';
import { Web } from '@pnp/pnpjs';
import chalk from 'chalk';
import fs from 'fs';

// Scan a site collection
export class Scan {
  static async run (config: ICrawlOptions): Promise<void> {
    // Crawled sites
    const sites: ISite[] = [];

    // Crawl counters
    const counters = {
      contentTypes: 0,
      fields: 0,
      sites: config.sites.length,
      webs: 0,
    };

    // Crawl sites
    for (const siteUrl of config.sites) {
      console.log(`Crawling webs in ${siteUrl}`);

      // Site
      const site: ISite = {
        LastScan: new Date().toISOString(),
        Url: siteUrl,
        Webs: [],
      };

      // Connect
      await new PnpNode().init().then(async settings => {
        // Initial webs
        const webUrls: string[] = [siteUrl];

        // Recursively iterate webs
        while (webUrls.length) {
          // Get next web
          const webUrl: string = _.trim(webUrls.shift(), '/');
          console.log(`  - ${chalk.whiteBright.bold(webUrl)}`);

          // Get web
          const web: Web = new Web(webUrl);

          // Get web data
          const webResponse = await web.select(...Object.keys(NewWeb)).get().catch(console.error);
          let webData = _.cloneDeep(NewWeb);
          for (const prop of Object.keys(NewWeb)) {
            if (webResponse[prop] !== undefined) {
              webData[prop] = webResponse[prop];
            }
          }
          site.Webs.push(webData);

          // Get content types
          const ctypes = await web.contentTypes.filter("Group ne '_Hidden'").select(...Object.keys(NewContentType)).get().catch(console.error) as any[];
          for (const ctype of ctypes) {
            // Add content type data
            let ctypeData = _.clone(NewContentType);
            for (const prop of Object.keys(NewContentType)) {
              if (ctype[prop] !== undefined) {
                if (ctype[prop] && ctype[prop].StringValue) ctypeData[prop] = ctype[prop].StringValue;
                else ctypeData[prop] = ctype[prop];
              }
            }
            webData.ContentTypes.push(ctypeData);

            // Increment content type counter
            counters.contentTypes++;
          }

          // Get fields
          const fields = await web.fields.filter("Hidden eq false").select(...Object.keys(NewField)).get().catch(console.error) as any[];
          for (const field of fields) {
            // Add field data
            let fieldData = _.clone(NewField);
            for (const prop of Object.keys(NewField)) {
              if (field[prop] !== undefined) {
                if (field[prop] && field[prop].StringValue) fieldData[prop] = field[prop].StringValue;
                else fieldData[prop] = field[prop];
              }
            }
            webData.Fields.push(fieldData);

            // Increment field counter
            counters.fields++;
          }

          // Get subwebs
          const subwebs = await web.webs.get().catch(console.error) as any[];

          // Add subwebs to webs array
          for (const subweb of subwebs) {
            webUrls.push(subweb.Url);
          }

          // Increment webs counter
          counters.webs++;
        }
      }).catch(console.error);

      // Add site
      sites.push(site);
    }

    // Write to file
    console.log(`Found ${counters.webs} webs, ${counters.contentTypes} content types, and ${counters.fields} fields in ${counters.sites} sites.\n  ${chalk.greenBright(Utility.path('webs.json'))}`);
    try {
      fs.writeFileSync(Utility.path('webs.json'), JSON.stringify(sites, null, 2), { flag: 'w' });
    } catch(e) {
      console.error(e);
    }
  }
}
