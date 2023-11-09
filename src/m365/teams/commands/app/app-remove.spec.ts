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
import command from './app-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let requests: any[];
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
    requests = [];
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      Cli.prompt,
      Cli.handleMultipleResultsFound,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'e3e29acb-8c79-412b-b746-e6c39ff4cd22',
        name: 'TeamsApp'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name options are not passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the id is not a valid GUID.', async () => {
    const actual = await command.validate({
      options: { id: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('removes Teams app by id in the tenant app catalog with confirmation (debug)',
    async () => {
      let removeTeamsAppCalled = false;
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
          removeTeamsAppCalled = true;
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, force: true } });
      assert(removeTeamsAppCalled);
    }
  );

  it('removes Teams app by id in the tenant app catalog without confirmation',
    async () => {
      let removeTeamsAppCalled = false;
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
          removeTeamsAppCalled = true;
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
      assert(removeTeamsAppCalled);
    }
  );

  it('aborts removing Teams app when prompt not confirmed', async () => {
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

    command.action(logger, { options: { id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } });
    assert(requests.length === 0);
  });

  it('removes Teams app by name in the tenant app catalog without confirmation (debug)',
    async () => {
      let removeTeamsAppCalled = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
          return {
            "value": [
              {
                "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
                "displayName": "TeamsApp"
              }
            ]
          };
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
          removeTeamsAppCalled = true;
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { debug: true, name: 'TeamsApp' } });
      assert(removeTeamsAppCalled);
    }
  );

  it('handles selecting single result when multiple teams apps to remove with the specified name are found and cli is set to prompt',
    async () => {
      let removeTeamsAppCalled = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
          return {
            "value": [
              { "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22" },
              { "id": "9b1b1e42-794b-4c71-93ac-5ed92488b67g" }
            ]
          };
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
          removeTeamsAppCalled = true;
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({ id: 'e3e29acb-8c79-412b-b746-e6c39ff4cd22' });
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { debug: true, name: 'TeamsApp' } });
      assert(removeTeamsAppCalled);
    }
  );

  it('fails to get Teams app when app does not exists', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'TeamsApp',
        force: true
      }
    } as any), new CommandError('The specified Teams app does not exist'));
  });

  it('handles error when multiple Teams apps with the specified name found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=displayName eq 'TeamsApp'&$select=id`) {
          return {
            "value": [
              {
                "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
                "displayName": "TeamsApp"
              },
              {
                "id": "5b31c38c-2584-42f0-aa47-657fb3a84230",
                "displayName": "TeamsApp"
              }
            ]
          };
        }
        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          name: 'TeamsApp',
          force: true
        }
      } as any), new CommandError(`Multiple Teams apps with name 'TeamsApp' found. Found: e3e29acb-8c79-412b-b746-e6c39ff4cd22, 5b31c38c-2584-42f0-aa47-657fb3a84230.`));
    }
  );

  it('correctly handles error when removing app', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects({
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    });
    await assert.rejects(command.action(logger, {
      options: {
        filePath: 'teamsapp.zip',
        id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
