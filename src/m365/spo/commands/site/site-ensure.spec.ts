import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import { WebProperties } from '../web/WebProperties.js';
import command from './site-ensure.js';

describe(commands.SITE_ENSURE, () => {
  let log: any[];
  let logger: Logger;
  const webResponse: WebProperties = {
    AllowRssFeeds: true,
    AlternateCssUrl: '',
    AppInstanceId: '00000000-0000-0000-0000-000000000000',
    Configuration: 0,
    Created: '2021-01-22T18:39:51.06',
    CurrentChangeToken: {
      StringValue: '1;2;113ba5b6-c737-4a6b-b1c7-2a367290057e;637470248884630000;125942029'
    },
    CustomMasterUrl: '/sites/team1/_catalogs/masterpage/seattle.master',
    Description: 'Team 2',
    DesignPackageId: '00000000-0000-0000-0000-000000000000',
    DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
    EnableMinimalDownload: false,
    HorizontalQuickLaunch: false,
    Id: '113ba5b6-c737-4a6b-b1c7-2a367290057e',
    IsMultilingual: true,
    Language: 1033,
    LastItemModifiedDate: '2021-01-22T18:44:16Z',
    LastItemUserModifiedDate: '2021-01-22T18:39:57Z',
    MasterUrl: '/sites/team1/_catalogs/masterpage/seattle.master',
    NoCrawl: false,
    OverwriteTranslationsOnChange: false,
    ResourcePath: {
      DecodedUrl: 'https://contoso.sharepoint.com/sites/team1'
    },
    QuickLaunchEnabled: true,
    RecycleBinEnabled: true,
    ServerRelativeUrl: '/sites/team1',
    SiteLogoUrl: '',
    SyndicationEnabled: true,
    Title: 'Team 2 updated',
    TreeViewEnabled: false,
    UIVersion: 15,
    UIVersionConfigurationEnabled: false,
    Url: 'https://contoso.sharepoint.com/sites/team1',
    WebTemplate: 'GROUP',
    AssociatedMemberGroup: '',
    AssociatedOwnerGroup: '',
    AssociatedVisitorGroup: ''
  };

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
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
  });

  afterEach(() => {
    jestUtil.restore([
      spo.getWeb,
      spo.addSite,
      spo.updateSite
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_ENSURE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates modern team site if no site found', async () => {
    jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().rejects({ error: '404 FILE NOT FOUND' });
    jest.spyOn(spo, 'addSite').mockClear().mockImplementation().resolves();

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1' } } as any);
  });

  it('creates modern communication site if no site found (debug)',
    async () => {
      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().rejects({ error: '404 FILE NOT FOUND' });
      jest.spyOn(spo, 'addSite').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/comms', title: 'Comms', type: 'CommunicationSite', debug: true } } as any);
    }
  );

  it('updates modern team site if existing modern team site found (no type specified)',
    async () => {
      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponse);
      jest.spyOn(spo, 'updateSite').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1' } } as any);
    }
  );

  it('updates modern team site if existing modern team site found (type specified)',
    async () => {
      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponse);
      jest.spyOn(spo, 'updateSite').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1', type: 'TeamSite' } } as any);
    }
  );

  it('updates modern communication site if existing modern communication site found (no type specified; debug)',
    async () => {
      const webResponseSitePagePublishing = { ...webResponse };
      webResponseSitePagePublishing.WebTemplate = 'SITEPAGEPUBLISHING';

      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponseSitePagePublishing);
      jest.spyOn(spo, 'updateSite').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/commsite1', title: 'CommSite1', debug: true } } as any);
    }
  );

  it('updates modern communication site if existing modern communication site found (type specified)',
    async () => {
      const webResponseSitePagePublishing = { ...webResponse };
      webResponseSitePagePublishing.WebTemplate = 'SITEPAGEPUBLISHING';

      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponseSitePagePublishing);
      jest.spyOn(spo, 'updateSite').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/commsite1', title: 'CommSite1', type: 'CommunicationSite' } } as any);
    }
  );

  it('updates classic site if an existing classic site found (type specified)',
    async () => {
      const webResponseSts = { ...webResponse };
      webResponseSts.WebTemplate = 'STS';

      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponseSts);
      jest.spyOn(spo, 'updateSite').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/classic', title: 'Classic', type: 'ClassicSite' } } as any);
    }
  );

  it(`updates site's visibility and sharing options`, async () => {
    jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponse);
    jest.spyOn(spo, 'updateSite').mockClear().mockImplementation().resolves();

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', alias: 'team1', title: 'Team 1', isPublic: true, shareByEmailEnabled: true } } as any);
  });

  it('returns error when validation of options for creating site failed',
    async () => {
      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().rejects(new Error('404 FILE NOT FOUND'));

      await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1' } } as any));
    }
  );

  it('returns error when an error has occurred when checking if a site exists at the specified URL',
    async () => {
      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

      await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1' } } as any));
    }
  );

  it('returns error when the specified site type is invalid', async () => {
    const webResponseSts = { ...webResponse };
    webResponseSts.WebTemplate = 'STS';

    jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponseSts);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/classic', title: 'Classic', type: 'Invalid' } } as any));
  });

  it('returns error when a communication site expected but a team site found',
    async () => {
      jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponse);

      await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1', title: 'Team 1', type: 'CommunicationSite' } } as any));
    }
  );

  it('returns error when no properties to update specified', async () => {
    jest.spyOn(spo, 'getWeb').mockClear().mockImplementation().resolves(webResponse);

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/team1' } } as any));
  });
});
