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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './hubsite-register.js';

describe(commands.HUBSITE_REGISTER, () => {
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_REGISTER);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('registers site as a hub site', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/RegisterHubSite`) > -1) {
        return {
          "Description": null,
          "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "LogoUrl": "http://contoso.com/logo.png",
          "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Test site"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "LogoUrl": "http://contoso.com/logo.png",
      "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
      "Targets": null,
      "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
      "Title": "Test site"
    }));
  });

  it('registers site as a hub site (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/site/RegisterHubSite`) > -1) {
        return {
          "Description": null,
          "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "LogoUrl": "http://contoso.com/logo.png",
          "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
          "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Test site"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/sales' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "ID": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "LogoUrl": "http://contoso.com/logo.png",
      "SiteId": "255a50b2-527f-4413-8485-57f4c17a24d1",
      "SiteUrl": "https://contoso.sharepoint.com/sites/sales",
      "Targets": null,
      "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
      "Title": "Test site"
    }));
  });

  it('correctly handles error when trying to register site which already is a hub site as a hub site',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
        error: {
          "odata.error": {
            "code": "-1, System.InvalidOperationException",
            "message": {
              "lang": "en-US",
              "value": "This site is already a HubSite."
            }
          }
        }
      });

      await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } } as any),
        new CommandError('This site is already a HubSite.'));
    }
  );

  it('fails validation if the specified site collection URL is not a valid SharePoint URL',
    async () => {
      const actual = await command.validate({ options: { siteUrl: 'site.com' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
