const _ = require('lodash');
const path = require('path');
const utility = require('./utility');
const { Web } = require('@pnp/sp');
const { PnpNode } = require('sp-pnp-node');

// Pull latest from git repository
module.exports = async function (config) {
  // Item counter
  let itemCount = 0;

  // URL array
  const urls = [];

  // Connect
  await new PnpNode().init().then(async settings => {
    // Initial webs array
    const webs = [settings.siteUrl];

    // Recursively iterate webs
    while (webs.length) {
      // Get next web
      const webUrl = _.trim(webs.shift(), '/');
      console.log(`Crawling ${webUrl}...`);

      // Get web
      const web = new Web(webUrl);

      // Get subwebs
      const subwebs = await web.webs.get();

      // Add subwebs to webs array
      for (const subweb of subwebs) {
        webs.push(subweb.Url);
      }

      // Get lists
      const lists = await web.lists.select(['Id', 'BaseType', 'Title', 'ItemCount', 'EntityTypeName']).filter('Hidden eq false').get();

      // Iterate lists
      for (const list of lists) {
        // Add list url URL
        const listUrl = `${webUrl}/${list.BaseType === 0 ? 'Lists/' : ''}${list.EntityTypeName}`;
        urls.push(listUrl);

        // Add items
        console.log(`  ${list.Title}: ${list.ItemCount} items (${listUrl})`);
        itemCount += list.ItemCount;

        // No items in library, or this is a list
        if (!list.BaseType || !list.ItemCount) continue;

        // Get list items
        const items = await web.lists.getById(list.Id).items.select(['Id', 'Title', 'FileRef']).getAll();

        // Save item URLs
        for (const item of items) {
          console.log(`    ${item.FileRef}`);
          urls.push(item.FileRef);
        }
      }
    }
  }).catch(console.error);

  // Done
  console.log(`Crawl complete. Found ${urls.length} URLs and ${itemCount} items in site collection.`);
  console.log(urls);
};
