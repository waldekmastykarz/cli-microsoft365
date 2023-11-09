import assert from 'assert';
import request from '../../../../request.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
import { jestUtil } from '../../../../utils/jestUtil.js';

class MockCommand extends SpoAppBaseCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('SpoAppBaseCommand', () => {
  const authSiteUrl = 'https://contoso.sharepoint.com';
  const tenantAppCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';

  const cmd: any = new MockCommand();
  const logger = {
    log: () => { },
    logRaw: () => { },
    logToStderr: () => { }
  };

  afterEach(() => {
    jestUtil.restore([
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('returns site collection URL correctly', async () => {
    const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, {
      options: {
        appCatalogUrl: 'https://contoso.sharepoint.com/sites/project-x',
        appCatalogScope: 'sitecollection'
      }
    });

    assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/project-x');
  });

  it('returns site collection URL correctly when it has a trailing slash',
    async () => {
      const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, {
        options: {
          appCatalogUrl: 'https://contoso.sharepoint.com/sites/project-x/',
          appCatalogScope: 'sitecollection'
        }
      });

      assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/project-x');
    }
  );

  it('returns site collection URL correctly when site collection app catalog URL is specified',
    async () => {
      const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, {
        options: {
          appCatalogUrl: 'https://contoso.sharepoint.com/sites/project-x/AppCatalog',
          appCatalogScope: 'sitecollection'
        }
      });

      assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/project-x');
    }
  );

  it('returns site collection URL correctly when site collection app catalog URL is specified with trailing slash',
    async () => {
      const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, {
        options: {
          appCatalogUrl: 'https://contoso.sharepoint.com/sites/project-x/AppCatalog/',
          appCatalogScope: 'sitecollection'
        }
      });

      assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/project-x');
    }
  );

  it(`returns site collection URL correctly when site collection is called 'appcatalog'`,
    async () => {
      const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, {
        options: {
          appCatalogUrl: 'https://contoso.sharepoint.com/sites/AppCatalog',
          appCatalogScope: 'sitecollection'
        }
      });

      assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/AppCatalog');
    }
  );

  it(`returns site collection URL correctly when site collection is called 'appcatalog' and app catalog URL is specified with trailing slash`,
    async () => {
      const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, {
        options: {
          appCatalogUrl: 'https://contoso.sharepoint.com/sites/AppCatalog/AppCatalog/',
          appCatalogScope: 'sitecollection'
        }
      });

      assert.strictEqual(actual, 'https://contoso.sharepoint.com/sites/AppCatalog');
    }
  );

  it('returns tenant app catalog URL correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `${authSiteUrl}/_api/SP_TenantSettings_Current`) {
        return {
          CorporateCatalogUrl: tenantAppCatalogUrl
        };
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    const actual = await cmd.getAppCatalogSiteUrl(logger, authSiteUrl, { options: {} });
    assert.strictEqual(actual, tenantAppCatalogUrl);
  });

  it('throws error when there is no tenant app catalog', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `${authSiteUrl}/_api/SP_TenantSettings_Current`) {
        return {
          CorporateCatalogUrl: null
        };
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(cmd.getAppCatalogSiteUrl(logger, authSiteUrl, { options: {} }), new Error('Tenant app catalog is not configured.'));
  });
});