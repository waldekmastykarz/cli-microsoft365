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
import command from './feature-list.js';

describe(commands.FEATURE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FEATURE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves available features from site collection', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
        DisplayName: "TenantSitesList"
      },
      {
        DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
        DisplayName: "Ratings"
      }
    ]));
  });

  it('retrieves available features from site', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            },
            {
              DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              DisplayName: "Ratings"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
        DisplayName: "TenantSitesList"
      },
      {
        DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
        DisplayName: "Ratings"
      }
    ]));
  });

  it('retrieves available features from site (default) when no scope is entered',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
          throw 'Invalid request';
        }

        if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
          return {
            value: [
              {
                DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
                DisplayName: "TenantSitesList"
              },
              {
                DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
                DisplayName: "Ratings"
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          verbose: false,
          webUrl: 'https://contoso.sharepoint.com'
        }
      });
      assert(loggerLogSpy.calledWith([
        {
          DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
          DisplayName: "TenantSitesList"
        },
        {
          DefinitionId: "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
          DisplayName: "Ratings"
        }
      ]));
    }
  );

  it('returns all properties for output JSON', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "odata.editLink": "Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
              "DefinitionId": "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              "DisplayName": "TenantSitesList"
            },
            {
              "odata.type": "SP.Feature",
              "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "odata.editLink": "Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
              "DefinitionId": "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
              "DisplayName": "Ratings"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const options: any = {
      debug: true,
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site',
      output: 'json'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(
      [
        {
          "odata.type": "SP.Feature",
          "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
          "odata.editLink": "Site/Features/GetById(guid'3019c9b4-e371-438d-98f6-0a08c34d06eb')",
          "DefinitionId": "3019c9b4-e371-438d-98f6-0a08c34d06eb",
          "DisplayName": "TenantSitesList"
        },
        {
          "odata.type": "SP.Feature",
          "odata.id": "https://contoso.sharepoint.com/_api/Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
          "odata.editLink": "Site/Features/GetById(guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3')",
          "DefinitionId": "915c240e-a6cc-49b8-8b2c-0bff8b553ed3",
          "DisplayName": "Ratings"
        }
      ]));
  });

  it('correctly handles no features in site collection', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return JSON.stringify({ value: [] });
      }

      throw 'Invalid request';
    });

    const options: any = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(log.length, 0);
  });

  it('correctly handles no features in site', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return JSON.stringify({ value: [] });
      }

      throw 'Invalid request';
    });

    const options: any = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(log.length, 0);
  });

  it('correctly handles no features in site collection (verbose)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
          return JSON.stringify({ value: [] });
        }

        throw 'Invalid request';
      });

      const options: any = {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      };
      await command.action(logger, { options: options } as any);
      let correctLogStatement = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No activated Features found') > -1) {
          correctLogStatement = true;
        }
      });
      assert(correctLogStatement);
    }
  );

  it('correctly handles no features in site (verbose)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return JSON.stringify({ value: [] });
      }

      throw 'Invalid request';
    });

    const options: any = {
      verbose: true,
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };
    await command.action(logger, { options: options } as any);
    let correctLogStatement = false;
    log.forEach(l => {
      if (!l || typeof l !== 'string') {
        return;
      }

      if (l.indexOf('No activated Features found') > -1) {
        correctLogStatement = true;
      }
    });
    assert(correctLogStatement);
  });

  it('correctly handles web feature reject request', async () => {
    const err = 'Invalid web Features reject request';
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }
    }), new CommandError(err));
  });

  it('correctly handles site Features reject request', async () => {
    const err = 'Invalid site Features reject request';
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }
    }), new CommandError(err));
  });

  it('supports specifying scope', () => {
    const options = command.options;
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('retrieves all Web features', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              DefinitionId: "00bfea71-5932-4f9c-ad71-1557e5751100",
              DisplayName: "WebPageLibrary"
            }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/abc', scope: 'Web' } });
    assert(loggerLogSpy.calledWith([
      {
        DefinitionId: '00bfea71-5932-4f9c-ad71-1557e5751100',
        DisplayName: 'WebPageLibrary'
      }]
    ));
  });

  it('retrieves all site features', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/Features?$select=DisplayName,DefinitionId') > -1) {
        return {
          value: [
            {
              DefinitionId: "3019c9b4-e371-438d-98f6-0a08c34d06eb",
              DisplayName: "TenantSitesList"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/abc', scope: 'Site' } });
    assert(loggerLogSpy.calledWith([
      {
        DefinitionId: '3019c9b4-e371-438d-98f6-0a08c34d06eb',
        DisplayName: 'TenantSitesList'
      }]
    ));
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({
        options:
        {
          webUrl: 'foo'
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when the url options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url and scope options specified',
    async () => {
      const actual = await command.validate({
        options:
        {
          webUrl: "https://contoso.sharepoint.com",
          scope: "Site"
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('accepts scope to be Site', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: 'Site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Web', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: 'Web'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid string scope', async () => {
    const scope = 'foo';
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('rejects invalid scope value specified as number', async () => {
    const scope = 123;
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('doesn\'t fail validation if the optional scope option not specified',
    async () => {
      const actual = await command.validate(
        {
          options:
          {
            webUrl: "https://contoso.sharepoint.com"
          }
        }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
