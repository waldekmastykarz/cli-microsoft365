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
import command from './managementapp-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.MANAGEMENTAPP_ADD, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
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
      request.get,
      request.put,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MANAGEMENTAPP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles error when the app specified with the objectId not found',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=id eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=appId`) {
          return { value: [] };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await assert.rejects(command.action(logger, {
        options: {
          objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        }
      }), new CommandError(`No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
    }
  );

  it('handles error when the app with the specified the name not found',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=appId`) {
          return { value: [] };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await assert.rejects(command.action(logger, {
        options: {
          name: 'My app'
        }
      }), new CommandError(`No Azure AD application registration with name My app found`));
    }
  );

  it('handles error when multiple apps with the specified name found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=appId`) {
          return {
            value: [
              { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await assert.rejects(command.action(logger, {
        options: {
          name: 'My app'
        }
      }), new CommandError("Multiple Azure AD application registration with name 'My app' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g."));
    }
  );

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20Test%20App'&$select=appId`) {
          return {
            value: [
              { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({
        "id": "340a4aa3-1af6-43ac-87d8-189819003952",
        "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
        "createdDateTime": "2019-10-29T17:46:55Z",
        "displayName": "My Test App",
        "description": null
      });

      jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('providers/Microsoft.BusinessAppPlatform/adminApplications/9b1b1e42-794b-4c71-93ac-5ed92488b67f?api-version=2020-06-01') > -1) {
          return {
            "applicationId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f"
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          name: 'My Test App', debug: true
        }
      });
      const call: sinon.SinonSpyCall = loggerLogSpy.mock.lastCall;
      assert.strictEqual(call.mock.calls[0].applicationId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    }
  );

  it('handles error when retrieving information about app through appId failed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

      await assert.rejects(command.action(logger, {
        options: {
          objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        }
      }), new CommandError(`An error has occurred`));
    }
  );

  it('handles error when retrieving information about app through name failed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

      await assert.rejects(command.action(logger, {
        options: {
          name: 'My app'
        }
      }), new CommandError(`An error has occurred`));
    }
  );

  it('fails validation if appId and objectId specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId, nor name specified',
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

  it('fails validation if the objectId is not a valid guid', async () => {
    const actual = await command.validate({ options: { objectId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (objectId)',
    async () => {
      const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { name: 'My app' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('successfully registers app as managementapp when passing appId',
    async () => {
      jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('providers/Microsoft.BusinessAppPlatform/adminApplications/9b1b1e42-794b-4c71-93ac-5ed92488b67f?api-version=2020-06-01') > -1) {
          return {
            "applicationId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f"
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
        }
      });
      const call: sinon.SinonSpyCall = loggerLogSpy.mock.lastCall;
      assert.strictEqual(call.mock.calls[0].applicationId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    }
  );

  it('successfully registers app as managementapp when passing name ',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20Test%20App'&$select=appId`) {
          return {
            value: [
              {
                "id": "340a4aa3-1af6-43ac-87d8-189819003952",
                "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
                "createdDateTime": "2019-10-29T17:46:55Z",
                "displayName": "My Test App",
                "description": null
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('providers/Microsoft.BusinessAppPlatform/adminApplications/9b1b1e42-794b-4c71-93ac-5ed92488b67f?api-version=2020-06-01') > -1) {
          return {
            "applicationId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f"
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          name: 'My Test App', debug: true
        }
      });
      const call: sinon.SinonSpyCall = loggerLogSpy.mock.lastCall;
      assert.strictEqual(call.mock.calls[0].applicationId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    }
  );
});
