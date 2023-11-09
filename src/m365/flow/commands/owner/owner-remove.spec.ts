import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { settingsNames } from '../../../../settingsNames.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './owner-remove.js';

describe(commands.OWNER_REMOVE, () => {
  const environmentName = 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6';
  const flowName = '1c6ee23a-a835-44bc-a4f5-462b658efc12';
  const userId = 'bb46fcd8-2104-4f3d-982c-eb74677251c2';
  const userName = 'john@contoso.com';
  const groupId = '37a0264d-fea4-4e87-8e5e-e574ff878cf2';
  const groupName = 'Test Group';
  const requestUrlNoAdmin = `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${formatting.encodeQueryParameter(flowName)}/modifyPermissions?api-version=2016-11-01`;
  const requestUrlAdmin = `https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}/flows/${formatting.encodeQueryParameter(flowName)}/modifyPermissions?api-version=2016-11-01`;
  const requestBodyUser = { 'delete': [{ 'id': userId }] };
  const requestBodyGroup = { 'delete': [{ 'id': groupId }] };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  let cli: Cli;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    cli = Cli.getInstance();
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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      aadGroup.getGroupIdByDisplayName,
      aadUser.getUserIdByUpn,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound,
      Cli.prompt,
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OWNER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes owner from flow by userId', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === requestUrlNoAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, userId: userId, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBodyUser);
  });

  it('deletes owner from flow by userName', async () => {
    jest.spyOn(aadUser, 'getUserIdByUpn').mockClear().mockImplementation().resolves(userId);
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === requestUrlNoAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, userName: userName, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBodyUser);
  });

  it('deletes owner from flow by groupId as admin when prompt confirmed',
    async () => {
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === requestUrlAdmin) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, groupId: groupId, asAdmin: true } });
      assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBodyGroup);
    }
  );

  it('deletes owner from flow by groupName as admin', async () => {
    jest.spyOn(aadGroup, 'getGroupIdByDisplayName').mockClear().mockImplementation().resolves(groupId);
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === requestUrlAdmin) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, groupName: groupName, asAdmin: true, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBodyGroup);
  });

  it('handles error when multiple groups with the specified name found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        return 'Invalid Request';
      });

      jest.spyOn(request, 'post').mockClear().mockImplementation().rejects('POST request executed');

      await assert.rejects(command.action(logger, {
        options: {
          verbose: true, environmentName: environmentName, flowName: flowName, groupName: groupName, asAdmin: true, force: true
        }
      }), new CommandError(`Multiple groups with name 'Test Group' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
    }
  );

  it('handles selecting single result when multiple groups with the name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '37a0264d-fea4-4e87-8e5e-e574ff878cf2' }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({ id: '37a0264d-fea4-4e87-8e5e-e574ff878cf2' });

      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === requestUrlAdmin) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, environmentName: environmentName, flowName: flowName, groupName: groupName, asAdmin: true, force: true } });
      assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBodyGroup);
    }
  );

  it('throws error when no environment found', async () => {
    const error = {
      error: {
        code: 'EnvironmentAccessDenied',
        message: `Access to the environment '${environmentName}' is denied.`
      }
    };
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, { options: { environmentName: environmentName, flowName: flowName, userId: userId, force: true } } as any),
      new CommandError(error.error.message));
  });

  it('prompts before removing the specified owner from a flow when confirm option not passed',
    async () => {
      await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, useName: userName } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified owner from a flow when option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

      await command.action(logger, { options: { environmentName: environmentName, flowName: flowName, useName: userName } });
      assert(postSpy.notCalled);
    }
  );

  it('fails validation if flowName is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: 'invalid', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if username is not a valid user principal name',
    async () => {
      const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, userName: 'invalid' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if groupName passed', async () => {
    const actual = await command.validate({ options: { environmentName: environmentName, flowName: flowName, groupName: groupName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
