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
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './app-permission-remove.js';

describe(commands.APP_PERMISSION_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  const validEnvironmentName = 'Default-6a2903af-9c03-4c02-a50b-e7419599925b';
  const validAppName = '784670e6-199a-4993-ae13-4b6747a0cd5d';
  const validUserId = 'd2481133-e3ed-4add-836d-6e200969dd03';
  const validUserName = 'john.doe@contoso.com';
  const validGroupId = 'c6c4b4e0-cd72-4d64-8ec2-cfbd0388ec16';
  const validGroupName = 'CLI Group';
  const appPermissionRemoveResponse = { put: [] };
  const groupResponse = {
    "id": validGroupId,
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2017-11-29T03:27:05Z",
    "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
    "displayName": "Finance",
    "groupTypes": [
      "Unified"
    ],
    "mail": "finance@contoso.onmicrosoft.com",
    "mailEnabled": true,
    "mailNickname": "finance",
    "onPremisesLastSyncDateTime": null,
    "onPremisesProvisioningErrors": [],
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "proxyAddresses": [
      "SMTP:finance@contoso.onmicrosoft.com"
    ],
    "renewedDateTime": "2017-11-29T03:27:05Z",
    "securityEnabled": false,
    "visibility": "Public"
  };
  const tenantId = '174290ec-373f-4d4c-89ea-9801dad0acd9';

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: '123',
      accessToken: 'abc'
    };
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

    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });

    jest.spyOn(accessToken, 'getTenantIdFromAccessToken').mockClear().mockReturnValue(tenantId);
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((_, defaultValue) => defaultValue);
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue,
      Cli.prompt,
      aadUser.getUserIdByUpn,
      aadGroup.getGroupByDisplayName,
      accessToken.getTenantIdFromAccessToken
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PERMISSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if appName is not a GUID', async () => {
    const actual = await command.validate({ options: { appName: 'invalid', userId: validUserId, force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if appName is a valid GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userId: validUserId, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not a GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userId: 'invalid', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userId: validUserId, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userName: 'John Doe', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userName: validUserName, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if groupId is not a GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, groupId: 'invalid', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, groupId: validGroupId, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if asAdmin is used without environmentName',
    async () => {
      const actual = await command.validate({ options: { appName: validAppName, asAdmin: true, userId: validUserId, force: true } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if environmentName is used without asAdmin',
    async () => {
      const actual = await command.validate({ options: { appName: validAppName, environmentName: validEnvironmentName, userId: validUserId, force: true } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if environmentName is used with asAdmin', async () => {
    const actual = await command.validate({ options: { appName: validAppName, environmentName: validEnvironmentName, userId: validUserId, asAdmin: true, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified Microsoft Power App permission when force option not passed',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation().resolves();

      await command.action(logger, {
        options: {
          appName: validAppName,
          userId: validUserId
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('removes permissions for a Power App by using user ID', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      delete: [{ id: validUserId }]
    };

    await command.action(logger, { options: { verbose: true, appName: validAppName, userId: validUserId, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
  });

  it('removes the permissions for the Power App for everyone and asks for confirmation',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
          return appPermissionRemoveResponse;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      const requestBody = {
        delete: [{ id: `tenant-${tenantId}` }]
      };

      await command.action(logger, { options: { appName: validAppName, tenant: true } });
      assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
    }
  );

  it('removes permissions for a Power App by using group ID', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      delete: [{ id: validGroupId }]
    };

    await command.action(logger, { options: { verbose: true, appName: validAppName, groupId: validGroupId, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
  });

  it('removes permissions for a Power App for everyone', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      delete: [{ id: `tenant-${tenantId}` }]
    };

    await command.action(logger, { options: { verbose: true, appName: validAppName, tenant: true, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
  });

  it('removes permissions for a Power App by using UPN', async () => {
    jest.spyOn(aadUser, 'getUserIdByUpn').mockClear().mockImplementation().resolves(validUserId);

    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      delete: [{ id: validUserId }]
    };

    await command.action(logger, { options: { verbose: true, appName: validAppName, userName: validUserName, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
  });

  it('removes permissions for a Power App by using group name', async () => {
    jest.spyOn(aadGroup, 'getGroupByDisplayName').mockClear().mockImplementation().resolves(groupResponse);

    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      delete: [{ id: validGroupId }]
    };

    await command.action(logger, { options: { verbose: true, appName: validAppName, groupName: validGroupName, force: true } });
    assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
  });

  it('removes permissions for a Power App by using user ID and running as admin',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
          return appPermissionRemoveResponse;
        }

        throw 'Invalid request';
      });

      const requestBody = {
        delete: [{ id: validUserId }]
      };

      await command.action(logger, { options: { appName: validAppName, userId: validUserId, environmentName: validEnvironmentName, asAdmin: true, force: true } });
      assert.deepStrictEqual(postStub.mock.lastCall[0].data, requestBody);
    }
  );

  it('correctly handles API OData error', async () => {
    const errorMessage = `The specified user with user id ${validUserId} does not exist`;
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      error: {
        error: {
          message: errorMessage
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { appName: validAppName, userId: validUserId, force: true } } as any),
      new CommandError(errorMessage));
  });
});
