import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './app-deploy.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_DEPLOY, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let requests: any[];

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    requests = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.post,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_DEPLOY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deploys app in the tenant app catalog (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('contextinfo') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    try {
      await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
    finally {
      jestUtil.restore([
        request.post,
        request.get
      ]);
    }
  });

  it('deploys app in the tenant app catalog', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    try {
      await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
    finally {
      jestUtil.restore([
        request.post,
        request.get
      ]);
    }
  });

  it('deploys app in the sitecollection app catalog', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });


    try {
      await command.action(logger, { options: { appCatalogScope: 'sitecollection', id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
    finally {
      jestUtil.restore([
        request.post,
        request.get
      ]);
    }
  });

  it('deploys app specified using its name in the tenant app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
          return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
        }

        if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) > -1) {
          return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { name: 'solution.sppkg' } });
    }
  );

  it('deploys app specified using its name in the sitecollection app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
          return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
        }

        if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) > -1) {
          return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { appCatalogScope: 'sitecollection', name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com' } });
    }
  );

  it('deploys app specified using its name in the tenant app catalog (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
          return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
        }

        if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) > -1) {
          return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, name: 'solution.sppkg' } });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('deploys app specified using its name in the site app catalog (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
          return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
        }
        if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) > -1) {
          return { UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6' };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, name: 'solution.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('deploys app in the tenant app catalog skipping feature deployment when the skipFeatureDeployment flag provided',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      try {
        await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', skipFeatureDeployment: true } });
        let correctRequestIssued = false;
        requests.forEach(r => {
          if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
            r.headers.accept &&
            r.headers.accept.indexOf('application/json') === 0 &&
            JSON.stringify(r.data) === JSON.stringify({ 'skipFeatureDeployment': true })) {
            correctRequestIssued = true;
          }
        });
        assert(correctRequestIssued);
      }
      finally {
        jestUtil.restore([
          request.post,
          request.get
        ]);
      }
    }
  );

  it('deploys app in the site app catalog skipping feature deployment when the skipFeatureDeployment flag provided',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      try {
        await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', skipFeatureDeployment: true, appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } });
        let correctRequestIssued = false;
        requests.forEach(r => {
          if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
            r.headers.accept &&
            r.headers.accept.indexOf('application/json') === 0 &&
            JSON.stringify(r.data) === JSON.stringify({ 'skipFeatureDeployment': true })) {
            correctRequestIssued = true;
          }
        });
        assert(correctRequestIssued);
      }
      finally {
        jestUtil.restore([
          request.post,
          request.get
        ]);
      }
    }
  );

  it('deploys app in the tenant app catalog not skipping feature deployment when the skipFeatureDeployment flag not provided',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      try {
        await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
        let correctRequestIssued = false;
        requests.forEach(r => {
          if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
            r.headers.accept &&
            r.headers.accept.indexOf('application/json') === 0 &&
            JSON.stringify(r.data) === JSON.stringify({ 'skipFeatureDeployment': false })) {
            correctRequestIssued = true;
          }
        });
        assert(correctRequestIssued);
      }
      finally {
        jestUtil.restore([
          request.post,
          request.get
        ]);
      }
    }
  );

  it('deploys app in the specified tenant app catalog', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });



    try {
      await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
    finally {
      jestUtil.restore([
        request.post,
        request.get
      ]);
    }
  });

  it('deploys app in the specified site app catalog', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    try {
      await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
    finally {
      jestUtil.restore([
        request.post,
        request.get
      ]);
    }
  });

  it('correctly deploys the app with valid URL provided in the prompt for tenant app catalog URL',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { appCatalogUrl: 'https://contoso.sharepoint.com' }
      ));

      try {
        await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
        let correctRequestIssued = false;
        requests.forEach(r => {
          if (r.url.indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1 &&
            r.headers.accept &&
            r.headers.accept.indexOf('application/json') === 0) {
            correctRequestIssued = true;
          }
        });
        assert(correctRequestIssued);
      }
      finally {
        jestUtil.restore([
          request.post,
          request.get
        ]);
      }
    }
  );

  it('correctly handles failure when app not found in app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return Promise.reject({
              error: {
                'odata.error': {
                  code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                  message: {
                    lang: "en-US",
                    value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                  }
                }
              }
            });
          }
        }

        throw 'Invalid request';
      });


      try {
        await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
          new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
      }
      finally {
        jestUtil.restore(request.post);
      }
    }
  );

  it('correctly handles failure when app not found in site app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return Promise.reject({
              error: {
                'odata.error': {
                  code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                  message: {
                    lang: "en-US",
                    value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                  }
                }
              }
            });
          }
        }

        throw 'Invalid request';
      });


      try {
        await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
          new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."));
      }
      finally {
        jestUtil.restore(request.post);
      }
    }
  );

  it('correctly handles failure when app specified by its name not found in app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
          return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
        }
        if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) > -1) {
          throw {
            error: {
              "odata.error": {
                "code": "-2147024894, System.IO.FileNotFoundException",
                "message": {
                  "lang": "en-US",
                  "value": "File Not Found."
                }
              }
            }
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/apps' } } as any),
        new CommandError('File Not Found.'));
    }
  );

  it('correctly handles failure when app specified by its name not found in site app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('SP_TenantSettings_Current') > -1) {
          return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" };
        }
        if ((opts.url as string).indexOf(`/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('solution.sppkg')?$select=UniqueId`) > -1) {
          throw {
            error: {
              "odata.error": {
                "code": "-2147024894, System.IO.FileNotFoundException",
                "message": {
                  "lang": "en-US",
                  "value": "File Not Found."
                }
              }
            }
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { name: 'solution.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('File Not Found.'));
    }
  );

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: 'An error has occurred' };
        }
      }

      throw 'Invalid request';
    });

    try {
      await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      jestUtil.restore(request.post);
    }
  });

  it('correctly handles random API error when site app catalog', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw { error: 'An error has occurred' };
        }
      }

      throw 'Invalid request';
    });


    try {
      await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      jestUtil.restore(request.post);
    }
  });

  it('correctly handles random API error (error message is not ODataError)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {

        if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            throw { error: JSON.stringify({ message: 'An error has occurred' }) };
          }
        }

        throw 'Invalid request';
      });

      try {
        await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
          new CommandError('{"message":"An error has occurred"}'));
      }
      finally {
        jestUtil.restore(request.post);
      }
    }
  );

  it('correctly handles random API error (error message is not ODataError) when site app catalog',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            throw { error: JSON.stringify({ message: 'An error has occurred' }) };
          }
        }

        throw 'Invalid request';
      });

      try {
        await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
          new CommandError('{"message":"An error has occurred"}'));
      }
      finally {
        jestUtil.restore(request.post);
      }
    }
  );

  it('correctly handles API OData error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
                message: {
                  value: 'An error has occurred'
                }
              }
            }
          };
        }
      }

      throw 'Invalid request';
    });

    try {
      await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      jestUtil.restore(request.post);
    }
  });

  it('correctly handles API OData error when scope is sitecollection',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/_api/web/sitecollectionappcatalog/AvailableApps/GetById('b2307a39-e878-458b-bc90-03bc578531d6')/deploy`) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            throw {
              error: {
                'odata.error': {
                  code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
                  message: {
                    value: 'An error has occurred'
                  }
                }
              }
            };
          }
        }

        throw 'Invalid request';
      });

      try {
        await assert.rejects(command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
          new CommandError('An error has occurred'));
      }
      finally {
        jestUtil.restore(request.post);
      }
    }
  );

  it('fails validation if neither the id nor the name are specified',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: {} }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both the id and the name are specified',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', name: 'solution.sppkg' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appCatalogUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation when the scope is specified invalid option',
    async () => {
      const actual = await command.validate({ options: { name: 'solution', appCatalogScope: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified',
    async () => {
      const actual = await command.validate({ options: { name: 'solution', appCatalogScope: 'sitecollection' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('should pass when \'tenant\' scope and also appCatalogUrl specified',
    async () => {
      const actual = await command.validate({ options: { name: 'solution', appCatalogScope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('should fail when \'sitecollection\' scope, but  bad appCatalogUrl format specified',
    async () => {
      const actual = await command.validate({ options: { name: 'solution', appCatalogScope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when the id is specified and the appCatalogUrl is not',
    async () => {
      const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when the id and appCatalogUrl options are specified',
    async () => {
      const actual = await command.validate({ options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'tenant' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when the name is specified and the appCatalogUrl is not',
    async () => {
      const actual = await command.validate({ options: { name: 'solution.sppkg' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when the name and appCatalogUrl options are specified',
    async () => {
      const actual = await command.validate({ options: { name: 'solution.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com', appCatalogScope: 'tenant' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when the name is specified without the extension',
    async () => {
      const actual = await command.validate({ options: { name: 'solution' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when the scope is specified with \'sitecollection\'',
    async () => {
      const actual = await command.validate({ options: { name: 'solution', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});