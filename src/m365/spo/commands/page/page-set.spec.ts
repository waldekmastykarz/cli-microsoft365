import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import spoFileGetCommand from '../file/file-get.js';
import spoListItemSetCommand from '../listitem/listitem-set.js';
import command from './page-set.js';

describe(commands.PAGE_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(spo, 'getRequestDigest').mockClear().mockImplementation().resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/article.aspx')/ListItemAllFields`) > -1 ||
        (opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sitepages/article.aspx')/ListItemAllFields`) > -1) &&
        JSON.stringify(opts.data) === JSON.stringify({
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/article.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.post,
      Cli.executeCommand,
      Cli.executeCommandWithOutput
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates page layout to Article', async () => {
    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      throw 'Invalid request';
    });
    jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: false, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Article (debug)', async () => {
    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      throw 'Invalid request';
    });
    jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } });
  });

  it('updates page layout to Article on root of tenant (debug)', async () => {
    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      throw 'Invalid request';
    });
    jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } });
  });

  it('automatically appends the .aspx extension', async () => {
    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      throw 'Invalid request';
    });
    jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: false, name: 'article', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Home', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields") {
        return;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home' } }));
    assert(loggerLogSpy.notCalled);
  });

  it('promotes the page as NewsPage', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      throw 'Invalid request';
    });

    jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        opts.data.PromotedState === 2 &&
        opts.data.FirstPublishedDate) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'NewsPage' } });
    assert(loggerLogSpy.notCalled);
  });

  it('promotes the page as Template', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      throw 'Invalid request';
    });

    jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        !opts.data) {
        return { Id: '1' };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(6)/SavePageAsTemplate`) > -1) {
        return { Id: '2', BannerImageUrl: 'url', CanvasContent1: 'content1', LayoutWebpartsContent: 'content' };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(2)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', description: "template", promoteAs: 'Template' } } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Home and promotes it as HomePage (debug)',
    async () => {
      jestUtil.restore([request.post]);
      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoFileGetCommand) {
          return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
        }
        throw 'Invalid request';
      });
      jest.spyOn(Cli, 'executeCommand').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoListItemSetCommand) {
          return;
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
          JSON.stringify(opts.data) === JSON.stringify({
            PageLayoutType: 'Home'
          })) {
          return;
        }

        if ((opts.url as string).indexOf('_api/web/rootfolder') > -1 &&
          opts.data.WelcomePage === 'SitePages/page.aspx') {
          return;
        }

        if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
          return {
            Title: "article",
            Id: 1,
            TopicHeader: "TopicHeader",
            AuthorByline: "AuthorByline",
            Description: "Description",
            BannerImageUrl: {
              Description: '/_layouts/15/images/sitepagethumbnail.png',
              Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
            },
            CanvasContent1: "{}",
            LayoutWebpartsContent: "{}"
          };
        }

        if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
          return;
        }

        if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home', promoteAs: 'HomePage', description: "Page Description" } });
    }
  );

  it('enables comments on the page', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: true } });
  });

  it('disables comments on the page (debug)', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields" ||
        opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(true)" ||
        opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(1)/SavePageAsDraft") {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: false } });
    assert(loggerLogSpy.notCalled);
  });

  it('demotes news page to a regular page', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoFileGetCommand) {
        return { 'stdout': '{\"FileSystemObjectType\":0,\"Id\":6,\"ServerRedirectedEmbedUri\":null,\"ServerRedirectedEmbedUrl\":\"\",\"ContentTypeId\":\"0x0101009D1CB255DA76424F860D91F20E6C411800E2DAFA6353688E488147257C551A63BD\",\"ComplianceAssetId\":null,\"WikiField\":null,\"Title\":\"zzzz\",\"CanvasContent1\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.0\\\" data-sp-controldata=\\\"&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;0&#125;&#125;\\\"><\/div><\/div>\",\"BannerImageUrl\":{\"Description\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\",\"Url\":\"https:\/\/contoso.sharepoint.com\/_layouts\/15\/images\/sitepagethumbnail.png\"},\"Description\":null,\"PromotedState\":0,\"FirstPublishedDate\":\"2022-11-11T15:48:15\",\"LayoutWebpartsContent\":\"<div><div data-sp-canvascontrol=\\\"\\\" data-sp-canvasdataversion=\\\"1.4\\\" data-sp-controldata=\\\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;audiences&quot;&#58;[],&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.4&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;zzzz&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;authors&quot;&#58;[&#123;&quot;id&quot;&#58;&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;,&quot;upn&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;email&quot;&#58;&quot;john.doe@contoso.onmicrosoft.com&quot;,&quot;name&quot;&#58;&quot;John Doe&quot;,&quot;role&quot;&#58;&quot;&quot;&#125;],&quot;authorByline&quot;&#58;[&quot;i&#58;0#.f|membership|john.doe@contoso.onmicrosoft.com&quot;]&#125;,&quot;reservedHeight&quot;&#58;228&#125;\\\"><\/div><\/div>\",\"OData__AuthorBylineId\":[9],\"_AuthorBylineStringId\":[\"9\"],\"OData__TopicHeader\":null,\"OData__SPSitePageFlags\":null,\"OData__SPCallToAction\":null,\"OData__OriginalSourceUrl\":null,\"OData__OriginalSourceSiteId\":null,\"OData__OriginalSourceWebId\":null,\"OData__OriginalSourceListId\":null,\"OData__OriginalSourceItemId\":null,\"ID\":6,\"Created\":\"2022-11-11T15:48:00\",\"AuthorId\":9,\"Modified\":\"2022-11-12T02:03:12\",\"EditorId\":9,\"OData__CopySource\":null,\"CheckoutUserId\":9,\"OData__UIVersionString\":\"2.19\",\"GUID\":\"9a94cb88-019b-4a66-abd6-be7f5337f659\"}' };
      }
      if (command === spoListItemSetCommand) {
        return;
      }
      throw 'Invalid request';
    });

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', demoteFrom: 'NewsPage' } });
  });

  it('updates page title', async () => {
    jestUtil.restore([request.post]);

    const newPageTitle = "updated title";

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: newPageTitle } });
  });

  it('updates page content', async () => {
    jestUtil.restore([request.post]);

    const newContent = [{
      "controlType": 4,
      "position": {
        "layoutIndex": 1,
        "zoneIndex": 1,
        "sectionIndex": 1,
        "sectionFactor": 12,
        "controlIndex": 1
      },
      "addedFromPersistedData": true,
      "innerHTML": "<p>Text content</p>"
    }];

    const initialPage = {
      Title: "article",
      Id: 1,
      TopicHeader: "TopicHeader",
      AuthorByline: "AuthorByline",
      Description: "Description",
      BannerImageUrl: {
        Description: '/_layouts/15/images/sitepagethumbnail.png',
        Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
      },
      CanvasContent1: "{}",
      LayoutWebpartsContent: "{}"
    };

    let data: string = '';
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).includes(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`)) {
        return initialPage;
      }
      if ((opts.url as string).includes(`/_api/SitePages/Pages(1)/SavePage`) ||
        (opts.url as string).includes(`/_api/SitePages/Pages(1)/SavePageAsDraft`)) {
        data = opts.data.CanvasContent1;
        return {};
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', content: JSON.stringify(newContent) } });
    assert.deepStrictEqual(JSON.parse(data), newContent);
  });

  it('publishes page', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1=\'\'&@a2=1`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('publishes page with a message (debug)', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1='Initial%20version'&@a2=1`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Initial version' } });
  });

  it('escapes special characters in user input', async () => {
    jestUtil.restore([request.post]);
    const comment = `Don't tell`;
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(comment)}'&@a2=1`) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: comment } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when creating modern page', async () => {
    jestUtil.restore([request.post]);

    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page layout', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--layoutType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page promote option', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--promoteAs') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if comments should be enabled', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--commentsEnabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if page should be published', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publish') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page publish message', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publishMessage') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'http://foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL',
    async () => {
      const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when name has no extension', async () => {
    const actual = await command.validate({ options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if layout type is invalid', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if layout type is Home', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Home' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Article', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is SingleWebPartAppPage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'SingleWebPartAppPage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is RepostPage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'RepostPage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is HeaderlessSearchResults',
    async () => {
      const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'HeaderlessSearchResults' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if layout type is Spaces', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Spaces' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Topic', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Topic' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if demote type is invalid', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', demoteFrom: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if promote type is invalid', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if promote type is HomePage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Home' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is NewsPage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is Template', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'Template' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is HomePage but layout type is not Home',
    async () => {
      const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Article' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if promote type is NewsPage but layout type is not Article',
    async () => {
      const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage', layoutType: 'Home' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if commentsEnabled is true', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if commentsEnabled is false', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if content is not valid JSON', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', content: "foo" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if content is valid JSON', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', content: "[]" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});