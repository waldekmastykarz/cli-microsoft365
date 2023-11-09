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
import command from './hubsite-unregister.js';

describe(commands.HUBSITE_UNREGISTER, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptOptions: any;

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
    requests = [];
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_UNREGISTER);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('unregisters the specified hub site without prompting with confirmation argument',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/site/UnregisterHubSite' &&
          opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales', force: true } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('prompts before unregistering the hub site when confirmation argument not passed',
    async () => {
      await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts unregistering hub site when prompt not confirmed', async () => {
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales' } });
    assert(requests.length === 0);
  });

  it('unregisters hub site when prompt confirmed', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/site/UnregisterHubSite' &&
        opts.headers &&
        opts.headers.accept &&
        (opts.headers.accept as string).indexOf('application/json') === 0) {
        return;
      }

      throw 'Invalid request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/sales' } });
  });

  it('correctly handles failure when the specified site is not a hub site',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/site/UnregisterHubSite') {
          throw {
            error: {
              "odata.error": {
                "code": "-2147024809, System.ArgumentException",
                "message": {
                  "lang": "en-US",
                  "value": "hubSiteId"
                }
              }
            }
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/sales', force: true } } as any),
        new CommandError("hubSiteId"));
    }
  );

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when the url is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
