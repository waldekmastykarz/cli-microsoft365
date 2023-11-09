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
import command from './hidedefaultthemes-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.HIDEDEFAULTTHEMES_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
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
    requests = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HIDEDEFAULTTHEMES_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets the value of the HideDefaultThemes setting', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('/_api/thememanager/SetHideDefaultThemes') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        hideDefaultThemes: true
      }
    });

    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/thememanager/SetHideDefaultThemes`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('sets the value of the HideDefaultThemes setting (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('/_api/thememanager/SetHideDefaultThemes') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        hideDefaultThemes: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/thememanager/SetHideDefaultThemes`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('handles error when setting the value of the HideDefaultThemes setting',
    async () => {
      const error = {
        error: {
          'odata.error': {
            code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      };

      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);
        if ((opts.url as string).indexOf('/_api/thememanager/SetHideDefaultThemes') > -1) {
          throw error;
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          hideDefaultThemes: true
        }
      } as any), new CommandError('An error has occurred'));
    }
  );

  it('fails validation if hideDefaultThemes is not set', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when hideDefaultThemes is true', async () => {
    const actual = await command.validate({ options: { hideDefaultThemes: true } }, commandInfo);
    assert(actual);
  });

  it('passes validation when hideDefaultThemes is false', async () => {
    const actual = await command.validate({ options: { hideDefaultThemes: false } }, commandInfo);
    assert(actual);
  });
});
