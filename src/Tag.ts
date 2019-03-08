import _ from 'lodash';
import { ICrawlOptions } from './Config';
import { PnpNode } from 'sp-pnp-node';
import { taxonomy } from '@pnp/sp-taxonomy';
import { Web, CheckinType } from '@pnp/pnpjs';
import chalk from 'chalk';

// Migrate data
const departments: any[] = [
  { name: 'Administration', id: 25 },
  { name: 'Sales, Marketing ＆ New Product Innovation', id: 14 },
  { name: 'Architecture ＆ Planning', id: 27 },
  { name: 'Assistant General Counsel', id: 19 },
  { name: 'C＆EN', id: 13 },
  { name: 'Membership ＆ Society Services', id: 6 },
  { name: 'Treasurer ＆ CFO', id: 22 },
  { name: 'Sales, Marketing ＆ New Product Innovation', id: 16 },
  { name: 'Publications', id: 45 },
  { name: 'Education', id: 32 },
  { name: 'Employee Services', id: 10 },
  { name: 'Education', id: 1 },
  { name: 'ACS Executive Leadership Team', id: 31 },
  { name: 'Finance', id: 23 },
  { name: 'Marketing ＆ Sales', id: 5 },
  { name: 'Green Chemistry Institute', id: 9 },
  { name: 'Education', id: 2 },
  { name: 'HRIS', id: 11 },
  { name: 'Human Resources', id: 35 },
  { name: 'Secretary ＆ General Counsel', id: 18 },
  { name: 'Meetings ＆ Exposition Services', id: 8 },
  { name: 'Investments', id: 21 },
  { name: 'Journals Publishing Group', id: 12 },
  { name: 'Education', id: 3 },
  { name: 'Meetings ＆ Exposition Services', id: 34 },
  { name: 'Member Programs ＆ Communities', id: 7 },
  { name: 'Membership ＆ Society Services', id: 33 },
  { name: 'External Affairs ＆ Communications', id: 20 },
  { name: 'Organization Development', id: 43 },
  { name: 'Learning and Career Development', id: 42 },
  { name: 'Learning and Career Development', id: 4 },
  { name: 'Publications Business Support', id: 26 },
  { name: 'Publications', id: 36 },
  { name: 'Research Grants', id: 46 },
  { name: 'Digital Strategy ＆ Publishing Platforms', id: 15 },
  { name: 'Scientific Advancement', id: 17 },
  { name: 'Secretary ＆ General Counsel', id: 37 },
  { name: 'Secretary ＆ General Counsel', id: 41 },
  { name: 'Society ＆ Administrative Technology', id: 29 },
  { name: 'Software ＆ Engineering', id: 28 },
  { name: 'Technical Infrastructure Services', id: 48 },
  { name: 'Technical Programs ＆ Activities', id: 47 },
  { name: 'Total Rewards', id: 44 },
  { name: 'Treasurer ＆ CFO', id: 38 },
  { name: 'Publications', id: 49 },
  { name: 'Washington IT Operations', id: 39 },
  { name: 'Web Strategy ＆ Operations', id: 30 },
];
const divisions: any[] = [
  { name: 'CAS', id: 1 },
  { name: 'Education', id: 2 },
  { name: 'ACS Executive Leadership Team', id: 10 },
  { name: 'Human Resources', id: 4 },
  { name: 'Membership ＆ Society Services', id: 3 },
  { name: 'Publications', id: 5 },
  { name: 'Scientific Advancement', id: 11 },
  { name: 'Secretary ＆ General Counsel', id: 7 },
  { name: 'Treasurer ＆ CFO', id: 8 },
  { name: 'Washington IT Operations', id: 9 },
];
const formCategories: any[] = [
  { name: 'Administration', id: 1 },
  { name: 'Finance', id: 2 },
  { name: 'Human Resources', id: 3 },
  { name: 'Washington IT Operations', id: 4 },
  { name: 'Meetings', id: 6 },
  { name: 'Membership', id: 5 },
  { name: 'Publications', id: 7 },
  { name: 'Travel', id: 8 },
];
const formTypes: any[] = [
  { name: 'Form', id: 1 },
  { name: 'Instructions', id: 2 },
  { name: 'Policy', id: 3 },
];
const offices: any[] = [
  { name: 'Accounts Payable', id: 'Accounts Payable' },
  { name: 'Benefits', id: 'Benefits' },
  { name: 'Budget ＆ Analysis', id: 'Budgets & Analysis' },
  { name: 'Conferencing ＆ Webinars', id: 'Conferencing' },
  { name: 'Contracts', id: 'Contract Administration' },
  { name: 'Finance', id: 'Finance' },
  { name: 'Financial Services', id: 'General Accounting' },
  { name: 'Mail, Print, ＆ Copy Services', id: 'Copy Center' },
  { name: 'National Meetings', id: 'National Meetings' },
  { name: 'Payroll', id: 'Payroll' },
  { name: 'Purchasing', id: 'Purchasing' },
  { name: 'Taxes', id: 'Taxes' },
];

// "Tags" term set ID
const termSetId = 'a664dbf2-f7d4-485e-820a-22757e1418e9';

// Migrate ACS
export class Tag {
  static async run (config: ICrawlOptions): Promise<void> {
    // Crawl sites
    for (const siteUrl of config.sites) {
      console.log(`Connecting to ${siteUrl}`);

      // Connect
      await new PnpNode().init().then(async settings => {
        console.log(`  Connected to ${settings.siteUrl}`);

        // Get term set
        console.log('Getting term store...');
        const store = await taxonomy.getDefaultSiteCollectionTermStore().get();
        const termset = await store.getTermSetById(termSetId).get();
        console.log(`  ${store.Name}/${termSetId}`);

        // Get all terms
        const terms = await termset.terms.get();
        console.log(`  ${terms.length} terms\n`);

        // Departments
        for (let i = 0; i < departments.length; i++) {
          // Find matching terms
          const matchingTerms = terms.filter(term => term.Name == departments[i].name);

          // Not found
          if (!matchingTerms.length) {
            console.log(`${chalk.bold(`${departments[i].name}:`)} ${chalk.red('not found')}`);
            continue;
          }

          // Update term
          let guid = matchingTerms[0].Id as string;
          guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
          departments[i].guid = guid;
          departments[i].term = matchingTerms[0];
          console.log(`${chalk.bold(`${departments[i].name}:`)} ${chalk.green(departments[i].guid)}`);
        }
        console.log('');

        // Divisions
        for (let i = 0; i < divisions.length; i++) {
          // Find matching terms
          const matchingTerms = terms.filter(term => term.Name == divisions[i].name);

          // Not found
          if (!matchingTerms.length) {
            console.log(`${chalk.bold(`${divisions[i].name}:`)} ${chalk.red('not found')}`);
            continue;
          }

          // Update term
          let guid = matchingTerms[0].Id as string;
          guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
          divisions[i].guid = guid;
          divisions[i].term = matchingTerms[0];
          console.log(`${chalk.bold(`${divisions[i].name}:`)} ${chalk.green(divisions[i].guid)}`);
        }
        console.log('');

        // Form categories
        for (let i = 0; i < formCategories.length; i++) {
          // Find matching terms
          const matchingTerms = terms.filter(term => term.Name == formCategories[i].name);

          // Not found
          if (!matchingTerms.length) {
            console.log(`${chalk.bold(`${formCategories[i].name}:`)} ${chalk.red('not found')}`);
            continue;
          }

          // Update term
          let guid = matchingTerms[0].Id as string;
          guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
          formCategories[i].guid = guid;
          formCategories[i].term = matchingTerms[0];
          console.log(`${chalk.bold(`${formCategories[i].name}:`)} ${chalk.green(formCategories[i].guid)}`);
        }
        console.log('');

        // Form categories
        for (let i = 0; i < formTypes.length; i++) {
          // Find matching terms
          const matchingTerms = terms.filter(term => term.Name == formTypes[i].name);

          // Not found
          if (!matchingTerms.length) {
            console.log(`${chalk.bold(`${formTypes[i].name}:`)} ${chalk.red('not found')}`);
            continue;
          }

          // Update term
          let guid = matchingTerms[0].Id as string;
          guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
          formTypes[i].guid = guid;
          formTypes[i].term = matchingTerms[0];
          console.log(`${chalk.bold(`${formTypes[i].name}:`)} ${chalk.green(formTypes[i].guid)}`);
        }
        console.log('');

        // Offices
        for (let i = 0; i < offices.length; i++) {
          // Find matching terms
          const matchingTerms = terms.filter(term => term.Name == offices[i].name);

          // Not found
          if (!matchingTerms.length) {
            console.log(`${chalk.bold(`${offices[i].name}:`)} ${chalk.red('not found')}`);
            continue;
          }

          // Update term
          let guid = matchingTerms[0].Id as string;
          guid = guid.replace(/\/Guid\((.*)\)\//ig, '$1');
          offices[i].guid = guid;
          offices[i].term = matchingTerms[0];
          console.log(`${chalk.bold(`${offices[i].name}:`)} ${chalk.green(offices[i].guid)}`);
        }
        console.log('');

        // Get web
        const web: Web = new Web(`${siteUrl}/policies-forms`);

        // Get "Tags" field
        const tagsField = await web.lists.getByTitle('Documents').fields.getByTitle('ACS_Tags_0').get();

        // Get all documents
        const documents = await web.lists.getByTitle('Documents').items.select('ID', 'Title', 'Migrate_Department', 'Migrate_Division', 'Migrate_FormCategory', 'Migrate_FormType', 'Migrate_Office', 'Keywords', 'File/Name', 'File/ServerRelativeUrl', 'File/Title', 'File/Length').expand('file').getAll().catch(e => console.error);

        // Iterate pages
        for (const document of documents as any[]) {
          if (!document.File || !document.File.ServerRelativeUrl) continue;
          console.log(`Updating ${document.File.ServerRelativeUrl}`);

          // Get location tag
          const locationTag = divisions[2];

          // Build tags
          const tags: any[] = [];

          // Department tags
          if (document.Migrate_Department) {
            const departmentId = parseInt(document.Migrate_Department);
            for (const department of departments) {
              if (department.id == departmentId && !tags.filter(tag => tag.guid == department.guid).length) {
                tags.push(department);
                break;
              }
            }
          }

          // Division tags
          if (document.Migrate_Division) {
            const divisionId = parseInt(document.Migrate_Division);
            for (const division of divisions) {
              if (division.id == divisionId && !tags.filter(tag => tag.guid == division.guid).length) {
                tags.push(division);
                break;
              }
            }
          }

          // Form category tags
          if (document.Migrate_FormCategory) {
            const formCategoryId = parseInt(document.Migrate_FormCategory);
            for (const formCategory of formCategories) {
              if (formCategory.id == formCategoryId && !tags.filter(tag => tag.guid == formCategory.guid).length) {
                tags.push(formCategory);
                break;
              }
            }
          }

          // Form type tags
          if (document.Migrate_FormType) {
            const formTypeId = parseInt(document.Migrate_FormType);
            for (const formType of formTypes) {
              if (formType.id == formTypeId && !tags.filter(tag => tag.guid == formType.guid).length) {
                tags.push(formType);
                break;
              }
            }
          }

          // Office tags
          if (document.Migrate_Office) {
            for (const office of offices) {
              if (office.id == document.Migrate_Office && !tags.filter(tag => tag.guid == office.guid).length) {
                tags.push(office);
                break;
              }
            }
          }

          // Add tags
          if (tags.length) {
            console.log(`  ${chalk.cyan(tags.map(tag => tag.name).join(', '))}`);

            // Build term string
            let termString = '';
            for (const tag of tags) {
              termString += `-1;#${tag.term.Name}|${tag.guid};#`;
            }

            // Update data
            const updateData: any = {};
            updateData[tagsField.InternalName] = termString;

            // Check out
            // await web.getFileByServerRelativeUrl(document.File.ServerRelativeUrl).checkout().catch(e => {});
            // console.log(chalk.gray('  Checked out'));

            // Update tags
            await web.lists.getByTitle('Documents').items.getById(document.ID).update(updateData).catch(e => console.error);
            console.log(chalk.gray('  Updated'));

            // Check in
            // const fileCheckedIn = await web.getFileByServerRelativeUrl(document.File.ServerRelativeUrl).checkin('Checked in via automated migration', CheckinType.Major).catch(e => {});
            // console.log(chalk.gray('  Checked in'));

            // // Approve
            // const fileApproved = await web.getFileByServerRelativeUrl(document.File.ServerRelativeUrl).approve('Approved via automated migration').catch(e => {});
            // console.log(chalk.gray('  Approved'));
          } else {
            console.log(chalk.gray('  No tags'));
          }
        }
      }).catch(console.error);
    }

    console.log('Done.');
  }
}
