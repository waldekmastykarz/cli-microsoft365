import assert from 'assert';
import fs from 'fs';
import auth from '../../../Auth.js';
import { CommandError } from '../../../Command.js';
import { Cli } from '../../../cli/Cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command from './app-open.js';

describe(commands.OPEN, () => {
  let log: string[];
  let logger: Logger;
  let cli: Cli;
  let openStub: jest.Mock;
  let getSettingWithDefaultValueStub: jest.Mock;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(JSON.stringify({
      "apps": [
        {
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "name": "CLI app1"
        }
      ]
    }));
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    cli = Cli.getInstance();
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
    openStub = jest.spyOn(browserUtil, 'open').mockClear().mockImplementation().resolves();
    getSettingWithDefaultValueStub = jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((() => false));
  });

  afterEach(() => {
    openStub.mockRestore();
    getSettingWithDefaultValueStub.mockRestore();
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OPEN);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the appId is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid appId-guid is specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows message with url when the app specified with the appId is found',
    async () => {
      const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
      await command.action(logger, {
        options: {
          appId: appId
        }
      });
      assert(loggerLogSpy.calledWith(`Use a web browser to open the page https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
    }
  );

  it('shows message with url when the app specified with the appId is found (verbose)',
    async () => {
      const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
      await command.action(logger, {
        options: {
          verbose: true,
          appId: appId
        }
      });
      assert(loggerLogSpy.calledWith(`Use a web browser to open the page https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
    }
  );

  it('shows message with preview-url when the app specified with the appId is found',
    async () => {
      const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
      await command.action(logger, {
        options: {
          appId: appId,
          preview: true
        }
      });
      assert(loggerLogSpy.calledWith(`Use a web browser to open the page https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
    }
  );

  it('shows message with url when the app specified with the appId is found (using autoOpenInBrowser)',
    async () => {
      getSettingWithDefaultValueStub.mockRestore();
      getSettingWithDefaultValueStub = jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockReturnValue(true);

      openStub.mockRestore();
      openStub = jest.spyOn(browserUtil, 'open').mockClear().mockImplementation(async (url) => {
        if (url === `https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`) {
          return;
        }
        throw 'Invalid url';
      });

      const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
      await command.action(logger, {
        options: {
          appId: appId
        }
      });
      assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
    }
  );

  it('shows message with preview-url when the app specified with the appId is found (using autoOpenInBrowser)',
    async () => {
      getSettingWithDefaultValueStub.mockRestore();
      getSettingWithDefaultValueStub = jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockReturnValue(true);

      openStub.mockRestore();
      openStub = jest.spyOn(browserUtil, 'open').mockClear().mockImplementation(async (url) => {
        if (url === `https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`) {
          return;
        }
        throw 'Invalid url';
      });

      const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
      await command.action(logger, {
        options: {
          appId: appId,
          preview: true
        }
      });
      assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
    }
  );

  it('throws error when open in browser fails', async () => {
    openStub.mockRestore();
    openStub = jest.spyOn(browserUtil, 'open').mockClear().mockImplementation(async (url) => {
      if (url === `https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`) {
        throw 'An error occurred';
      }
      throw 'Invalid url';
    });

    getSettingWithDefaultValueStub.mockRestore();
    getSettingWithDefaultValueStub = jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockReturnValue(true);

    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await assert.rejects(command.action(logger, {
      options: {
        appId: appId,
        preview: true
      }
    }), new CommandError('An error occurred'));
    assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });
});
