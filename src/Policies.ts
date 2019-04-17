import _ from 'lodash';
import { ICrawlOptions } from './Config';
import { PnpNode } from 'sp-pnp-node';
import { Web, CheckinType } from '@pnp/pnpjs';
import chalk from 'chalk';

// Migrate ACS policy pages
export class Policies {
  static async run (config: ICrawlOptions): Promise<void> {
    // Crawl sites
    for (const siteUrl of config.sites) {
      console.log(`Connecting to ${siteUrl}`);

      // Connect
      await new PnpNode().init().then(async settings => {
        console.log(`  Connected to ${settings.siteUrl}`);

        // Get web
        const web: Web = new Web(`${siteUrl}/policy-center`);

        // Get all pages
        const pages = await web.lists.getByTitle('Pages').items.select('*').expand('file').getAll().catch(e => console.error);

        // Iterate pages
        for (const page of pages as any[]) {
          if (!page.File || !page.File.ServerRelativeUrl || !page.ACSRelatedPolicies) continue;
          console.log(`Updating ${page.File.ServerRelativeUrl}`);

          // Check out
          await web.getFileByServerRelativeUrl(page.File.ServerRelativeUrl).checkout().catch(e => {});
          console.log('  Checked out');

          // Replace related policy URLs
          // let relatedPolicies = page.ACSRelatedPolicies.replace(/https:\/\/intranet\.acs\.org\/policycenter\//ig, '/policies-forms/policies-procedures/').replace(/\/policycenter\//ig, '/policies-forms/policies-procedures/');
          let relatedPolicies = page.ACSRelatedPolicies.replace(/\/policies-forms\/policies-procedures\//ig, '/policy-center/');

          // Replace related IDs
          // for (const rPage of pages as any[]) {
          //   const searchTitle = new RegExp(`"Id":"\\d+","Title":"${rPage.Title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'ig');
          //   relatedPolicies = relatedPolicies.replace(searchTitle, `"Id":"${rPage.ID}","Title":"${rPage.Title}"`);
          // }

          // Update related policies
          await web.lists.getByTitle('Pages').items.getById(page.ID).update({
            ACSRelatedPolicies: relatedPolicies,
          }).catch(e => console.error);
          console.log('  Updated');

          // Check in
          const fileCheckedIn = await web.getFileByServerRelativeUrl(page.File.ServerRelativeUrl).checkin('Checked in via automated migration', CheckinType.Major).catch(e => {});
          console.log('  Checked in');

          // Approve
          const fileApproved = await web.getFileByServerRelativeUrl(page.File.ServerRelativeUrl).approve('Approved via automated migration').catch(e => {});
          console.log('  Approved');
        }
      }).catch(console.error);
    }

    console.log('Done.');
  }
}
