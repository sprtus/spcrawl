import _ from 'lodash';
import { ICrawlOptions } from './Config';
import { PnpNode } from 'sp-pnp-node';
import { Utility } from './Utility';
import { Web, CheckinType } from '@pnp/pnpjs';
import chalk from 'chalk';
import fs from 'fs';

// Migrate ACS news
export class Phoenix {
  static async crawl (config: ICrawlOptions): Promise<void> {
    // Crawl sites
    for (const siteUrl of config.sites) {
      console.log(`Connecting to ${siteUrl}`);

      // Connect
      await new PnpNode().init().then(async settings => {
        console.log(`  Connected to ${settings.siteUrl}`);

        // News articles
        const articles: any[] = [];

        // Get web
        const web: Web = new Web(`${siteUrl}/phoenix`);

        // Get all documents
        console.log(`\nGetting documents`);
        const files = await web.getFolderByServerRelativeUrl(`/phoenix/The Phoenix Archives`).files.select('ServerRelativeUrl', 'Title', 'Name', 'TimeCreated', 'Month', 'Year').top(5000).get().catch(e => console.error);

        // Iterate pages
        for (const file of files as any[]) {
          if (!file.ServerRelativeUrl) continue;
          console.log(`  ${file.Name}`);

          // Parse date
          const fileName = file.Name.split('.')[0];
          const d: any = {
            Year: null,
            Month: null,
            MonthText: null,
          };
          let duplicate = false;
          if (fileName.split('_').length === 3) {
            d.Year = parseInt(fileName.split('_')[0]);
            d.Month = parseInt(fileName.split('_')[1]);
            d.MonthText = fileName.split('_')[2];
            if (d.MonthText.indexOf('-') !== -1) {
              d.MonthText = d.MonthText.split('-')[0];
              duplicate = true;
            }
          }

          // Create title
          const title = `The Phoenix: ${d.Year ? `${d.MonthText}, ${d.Year}` : fileName}${duplicate ? ' (2)' : ''}`;

          // Create page name
          const pageName = `The-Phoenix-${d.Year ? `${d.Year}-${d.Month}${duplicate ? '-2' : ''}` : fileName}.aspx`;

          // Save article
          articles.push({
            Title: title,
            PageName: pageName,
            Date: d,
            Duplicate: duplicate,
            Url: `${siteUrl}${file.ServerRelativeUrl}`,
            File: file,
          });
        }

        // Write to file
        console.log(`\nSaving ${articles.length} articles...\n  ${chalk.greenBright(Utility.path('phoenix.json'))}`);
        try {
          fs.writeFileSync(Utility.path('phoenix.json'), JSON.stringify(articles, null, 2), { flag: 'w' });
        } catch(e) {
          console.error(e);
        }
      });
    }

    console.log('Done.');
  }

  static async create (config: ICrawlOptions): Promise<void> {
    // News site
    const newsSiteUrl = 'https://nucleusdev.acs.org/news-events';
    console.log(`Connecting to ${newsSiteUrl}`);

    // Connect
    await new PnpNode({
      siteUrl: newsSiteUrl,
    }).init().then(async settings => {
      console.log(`  Connected to ${settings.siteUrl}\n`);

      // Get news articles
      let articles: any = fs.readFileSync(Utility.path('./phoenix.json'));
      articles = JSON.parse(articles);

      // Get web
      const web: Web = new Web(`${newsSiteUrl}`);

      // Create pages
      for (const article of articles) {
        // Get relative page path
        const relativePageUrl = `/news-events/Pages/${article.PageName}`;
        console.log(`Creating: ${chalk.cyan(relativePageUrl)}`);

        // Create page content
        const pageContent = `<%@ Page Inherits="Microsoft.SharePoint.Publishing.TemplateRedirectionPage,Microsoft.SharePoint.Publishing,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c" %> <%@ Reference VirtualPath="~TemplatePageUrl" %> <%@ Reference VirtualPath="~masterurl/custom.master" %><%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
        <html xmlns:mso="urn:schemas-microsoft-com:office:office" xmlns:msdt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"><head>
        <!--[if gte mso 9]><SharePoint:CTFieldRefs runat=server Prefix="mso:" FieldList="FileLeafRef,Comments,PublishingStartDate,PublishingExpirationDate,PublishingContactEmail,PublishingContactName,PublishingContactPicture,PublishingPageLayout,PublishingVariationGroupID,PublishingVariationRelationshipLinkFieldID,PublishingRollupImage,Audience,PublishingIsFurlPage,SeoBrowserTitle,SeoMetaDescription,SeoKeywords,RobotsNoIndex,ACS_Summary,PublishingPageContent,ACS_Featured,odea017de2994eadb29fca35aaab7db7,TaxCatchAllLabel,ACS_Thumbnail,ACS_ShowSideNav,ACS_DivisionLanding,ACS_TeamLanding,ACS_EntryDate,ACS_Byline,n53db4776dca48dba0972095b72739fe,ACS_DigestSend,ACS_DigestOrder,ACS_DigestPubStart,ACS_DigestPubEnd,ACS_WhoToCallListName"><xml>
        <mso:CustomDocumentProperties>
        <mso:PublishingContact msdt:dt="string">778</mso:PublishingContact>
        <mso:PublishingIsFurlPage msdt:dt="string">0</mso:PublishingIsFurlPage>
        <mso:display_urn_x003a_schemas-microsoft-com_x003a_office_x003a_office_x0023_PublishingContact msdt:dt="string">WFService, SharePoint</mso:display_urn_x003a_schemas-microsoft-com_x003a_office_x003a_office_x0023_PublishingContact>
        <mso:ACS_DigestSend msdt:dt="string">1</mso:ACS_DigestSend>
        <mso:PublishingContactPicture msdt:dt="string"></mso:PublishingContactPicture>
        <mso:RobotsNoIndex msdt:dt="string">0</mso:RobotsNoIndex>
        <mso:ACS_ShowSideNav msdt:dt="string">0</mso:ACS_ShowSideNav>
        <mso:ACS_DivisionLanding msdt:dt="string">0</mso:ACS_DivisionLanding>
        <mso:ACS_EntryDate msdt:dt="string">${article.Date.Year ? `${article.Date.Year}-${article.Date.Month}-01T00:00:00Z` : article.File.TimeCreated}</mso:ACS_EntryDate>
        <mso:ACS_Send msdt:dt="string">0</mso:ACS_Send>
        <mso:ContentTypeId msdt:dt="string">0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF390059178973DC71AF439E03FE69E2CBC20F02002748B97C4D70AD46A057697FC45223CB</mso:ContentTypeId>
        <mso:PublishingContactName msdt:dt="string"></mso:PublishingContactName>
        <mso:ACS_Featured msdt:dt="string">0</mso:ACS_Featured>
        <mso:PublishingPageLayoutName msdt:dt="string">News.aspx</mso:PublishingPageLayoutName>
        <mso:Comments msdt:dt="string"></mso:Comments>
        <mso:PublishingContactEmail msdt:dt="string"></mso:PublishingContactEmail>
        <mso:PublishingPageLayout msdt:dt="string">https://nucleusdev.acs.org/_catalogs/masterpage/acs-nucleus/News.aspx, News</mso:PublishingPageLayout>
        <mso:ACS_TeamLanding msdt:dt="string">0</mso:ACS_TeamLanding>
        <mso:PublishingPageContent msdt:dt="string">&lt;p&gt;&lt;strong&gt;The Phoenix&lt;/strong&gt;${article.Date.Year ? ` issue for ${article.Date.MonthText}, ${article.Date.Year.toString()}` : ''} is available. Click on the link below to view the newsletter.&lt;/p&gt;&lt;p class=&quot;ms-rteElement-ButtonRowPrimary&quot&gt;&lt;a href=&quot;${article.Url}&quot;&gt;[icon:far fa-file-pdf] ${article.File.Name}&lt;/a&gt;&lt;/p&gt;</mso:PublishingPageContent>
        <mso:ACS_DigestPubStart msdt:dt="string"></mso:ACS_DigestPubStart>
        <mso:ACS_DigestOrder msdt:dt="string"></mso:ACS_DigestOrder>
        <mso:ACS_Summary msdt:dt="string">The${article.Date.Year ? ` ${article.Date.MonthText}, ${article.Date.Year.toString()} edition of` : ''} The Phoenix</mso:ACS_Summary>
        <mso:ACS_DigestPubEnd msdt:dt="string"></mso:ACS_DigestPubEnd>
        <mso:Order msdt:dt="string"></mso:Order>
        <mso:ACS_Tags msdt:dt="string"></mso:ACS_Tags>
        <mso:odea017de2994eadb29fca35aaab7db7 msdt:dt="string"></mso:odea017de2994eadb29fca35aaab7db7>
        <mso:n53db4776dca48dba0972095b72739fe msdt:dt="string">The Phoenix|36b4dd5f-3571-44be-a225-6154c18f480f</mso:n53db4776dca48dba0972095b72739fe>
        <mso:ACS_NewsChannel msdt:dt="string">58;#The Phoenix|36b4dd5f-3571-44be-a225-6154c18f480f</mso:ACS_NewsChannel>
        <mso:TaxCatchAll msdt:dt="string">58;#The Phoenix|36b4dd5f-3571-44be-a225-6154c18f480f</mso:TaxCatchAll>
        </mso:CustomDocumentProperties>
        </xml></SharePoint:CTFieldRefs><![endif]-->
        <title>${article.Title}</title></head>`;

        // Check out
        const checkedOut = await web.getFileByServerRelativeUrl(relativePageUrl).checkout().catch(e => {});
        console.log(chalk.gray('  Checked out'));

        // Update file
        const page: any = await web.getFolderByServerRelativeUrl('/news-events/Pages').files.add(article.PageName, pageContent, true).catch(e => console.error);
        console.log(chalk.gray('  Updated'));

        // Check in
        const fileCheckedIn = await web.getFileByServerRelativeUrl(relativePageUrl).checkin('Checked in via automated migration', CheckinType.Major).catch(e => {});
        console.log(chalk.gray('  Checked in'));

        // Approve
        const fileApproved = await web.getFileByServerRelativeUrl(relativePageUrl).approve('Approved via automated migration').catch(e => {});
        console.log(chalk.gray('  Approved'));
      }
    });

    console.log('\nDone.');
  }
}
