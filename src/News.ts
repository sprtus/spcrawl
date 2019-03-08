import _ from 'lodash';
import { ICrawlOptions } from './Config';
import { PnpNode } from 'sp-pnp-node';
import { Utility } from './Utility';
import { Web, CheckinType } from '@pnp/pnpjs';
import chalk from 'chalk';
import fs from 'fs';

const months: any = {
  '01': 'January',
  '02': 'February',
  '03': 'March',
  '04': 'April',
  '05': 'May',
  '06': 'June',
  '07': 'July',
  '08': 'August',
  '09': 'September',
  '10': 'October',
  '11': 'November',
  '12': 'December',
};

// Migrate ACS news
export class News {
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
        const web: Web = new Web(`${siteUrl}/Publications/SalesMarketing/WebInnovation`);

        for (let year = 2006; year <= 2017; year++) {
          // Get all documents
          console.log(`\nGetting documents: ${chalk.cyan(year.toString())}`);
          const files = await web.getFolderByServerRelativeUrl(`/Publications/SalesMarketing/WebInnovation/StatusReports/Market Intelligence Report/Market Intelligence Report - ${year.toString()}`).files.select('ServerRelativeUrl', 'Title', 'Name', 'TimeCreated').top(5000).get().catch(e => console.error);

          // Iterate pages
          for (const file of files as any[]) {
            if (!file.ServerRelativeUrl) continue;

            // Create news title
            let title = file.Name.split('.')[0];
            let month = null;

            // Parse date
            if ((title.length == 7 && title.split('-').length == 2) || (title.length == 10 && title.split('-').length == 3)) {
              month = title.split('-')[1];
              if (months[month]) {
                month = months[month];
                title = `Market Intelligence Report: ${month}, ${year}`;
              }
            }

            // Save article
            console.log(`  ${file.Name}: ${chalk.green(title)}`);
            articles.push({
              Title: title,
              Url: `${siteUrl}${file.ServerRelativeUrl}`,
              Date: file.TimeCreated,
              Name: file.Name,
              Month: month,
              Year: year,
            });
          }
        }

        // Write to file
        console.log(`\nSaving ${articles.length} articles...\n  ${chalk.greenBright(Utility.path('news.json'))}`);
        try {
          fs.writeFileSync(Utility.path('news.json'), JSON.stringify(articles, null, 2), { flag: 'w' });
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
      let articles: any = fs.readFileSync(Utility.path('./news.json'));
      articles = JSON.parse(articles);

      // Get web
      const web: Web = new Web(`${newsSiteUrl}`);

      // Create pages
      for (const article of articles) {
        // Get relative page path
        const pageName = `${article.Title.replace(/:\s/g, '-').replace(/,\s/g, '-').replace(/\s/g, '-')}.aspx`;
        const relativePageUrl = `/news-events/Pages/${pageName}`;
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
        <mso:ACS_EntryDate msdt:dt="string">${article.Date}</mso:ACS_EntryDate>
        <mso:ACS_Send msdt:dt="string">0</mso:ACS_Send>
        <mso:ContentTypeId msdt:dt="string">0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF390059178973DC71AF439E03FE69E2CBC20F02002748B97C4D70AD46A057697FC45223CB</mso:ContentTypeId>
        <mso:PublishingContactName msdt:dt="string"></mso:PublishingContactName>
        <mso:ACS_Featured msdt:dt="string">0</mso:ACS_Featured>
        <mso:PublishingPageLayoutName msdt:dt="string">News.aspx</mso:PublishingPageLayoutName>
        <mso:Comments msdt:dt="string"></mso:Comments>
        <mso:PublishingContactEmail msdt:dt="string"></mso:PublishingContactEmail>
        <mso:PublishingPageLayout msdt:dt="string">https://nucleusdev.acs.org/_catalogs/masterpage/acs-nucleus/News.aspx, News</mso:PublishingPageLayout>
        <mso:ACS_TeamLanding msdt:dt="string">0</mso:ACS_TeamLanding>
        <mso:PublishingPageContent msdt:dt="string">&lt;p&gt;The latest &lt;strong&gt;Market Intelligence Reporter&lt;/strong&gt;${article.Month ? ` for ${article.Month}, ${article.Year.toString()}` : ''} is now available. Click on the link below to view the report.&lt;/p&gt;&lt;p class=&quot;ms-rteElement-ButtonRowPrimary&quot&gt;&lt;a href=&quot;${article.Url}&quot;&gt;[icon:far fa-file-pdf] ${article.Name}&lt;/a&gt;&lt;/p&gt;</mso:PublishingPageContent>
        <mso:ACS_DigestPubStart msdt:dt="string"></mso:ACS_DigestPubStart>
        <mso:ACS_DigestOrder msdt:dt="string"></mso:ACS_DigestOrder>
        <mso:ACS_Summary msdt:dt="string">The monthly Market Intelligence Reporter</mso:ACS_Summary>
        <mso:ACS_DigestPubEnd msdt:dt="string"></mso:ACS_DigestPubEnd>
        <mso:Order msdt:dt="string"></mso:Order>
        <mso:ACS_Tags msdt:dt="string"></mso:ACS_Tags>
        <mso:odea017de2994eadb29fca35aaab7db7 msdt:dt="string"></mso:odea017de2994eadb29fca35aaab7db7>
        <mso:n53db4776dca48dba0972095b72739fe msdt:dt="string">Market Intelligence Reporter|221d6439-169c-43d3-9b14-d5e8cb681b88</mso:n53db4776dca48dba0972095b72739fe>
        <mso:ACS_NewsChannel msdt:dt="string">58;#Market Intelligence Reporter|221d6439-169c-43d3-9b14-d5e8cb681b88</mso:ACS_NewsChannel>
        <mso:TaxCatchAll msdt:dt="string">58;#Market Intelligence Reporter|221d6439-169c-43d3-9b14-d5e8cb681b88</mso:TaxCatchAll>
        </mso:CustomDocumentProperties>
        </xml></SharePoint:CTFieldRefs><![endif]-->
        <title>${article.Title}</title></head>`;

        // Check out
        const checkedOut = await web.getFileByServerRelativeUrl(relativePageUrl).checkout().catch(e => {});
        console.log(chalk.gray('  Checked out'));

        // Update file
        const page: any = await web.getFolderByServerRelativeUrl('/news-events/Pages').files.add(pageName, pageContent, true).catch(e => console.error);
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
