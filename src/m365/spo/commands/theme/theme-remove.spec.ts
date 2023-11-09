import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './theme-remove.js';

describe(commands.THEME_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    promptOptions = undefined;
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should prompt before removing theme when confirmation argument not passed',
    async () => {
      await command.action(logger, { options: { name: 'Contoso' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('removes theme successfully without prompting with confirmation argument',
    async () => {
      const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {

        if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
          return 'Correct Url';
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          name: 'Contoso',
          force: true
        }
      });
      assert.strictEqual(postStub.mock.lastCall[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
      assert.strictEqual(postStub.mock.lastCall[0].headers['accept'], 'application/json;odata=nometadata');
      assert.strictEqual(postStub.mock.lastCall[0].data.name, 'Contoso');
      assert.strictEqual(loggerLogSpy.notCalled, true);
    }
  );

  it('removes theme successfully without prompting with confirmation argument (debug)',
    async () => {
      const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {

        if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
          return 'Correct Url';
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          name: 'Contoso',
          force: true
        }
      });
      assert.strictEqual(postStub.mock.lastCall[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
      assert.strictEqual(postStub.mock.lastCall[0].headers['accept'], 'application/json;odata=nometadata');
      assert.strictEqual(postStub.mock.lastCall[0].data.name, 'Contoso');
    }
  );

  it('removes theme successfully when prompt confirmed', async () => {
    const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso'
      }
    });
    assert.strictEqual(postStub.mock.lastCall[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
    assert.strictEqual(postStub.mock.lastCall[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(postStub.mock.lastCall[0].data.name, 'Contoso');
  });

  it('handles error when removing theme', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
