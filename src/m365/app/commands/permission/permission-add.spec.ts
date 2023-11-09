import { Application, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import fs from 'fs';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './permission-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.PERMISSION_ADD, () => {
  //#region Mocked responses
  const appId = '9c79078b-815e-4a3e-bb80-2aaf2d9e9b3d';
  const servicePrincipalId = '7c330108-8825-4b6c-b280-8d1d68da6bd7';
  const servicePrincipals: ServicePrincipal[] = [{ "appId": appId, 'id': servicePrincipalId, "servicePrincipalNames": [] }, { "appId": "00000003-0000-0000-c000-000000000000", "id": "fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be", "servicePrincipalNames": ["https://canary.graph.microsoft.com/", "https://graph.microsoft.us/", "https://dod-graph.microsoft.us/", "00000003-0000-0000-c000-000000000000/ags.windows.net", "00000003-0000-0000-c000-000000000000", "https://canary.graph.microsoft.com", "https://graph.microsoft.com", "https://ags.windows.net", "https://graph.microsoft.us", "https://graph.microsoft.com/", "https://dod-graph.microsoft.us"], "appRoles": [{ "allowedMemberTypes": ["Application"], "description": "Allows the app to read and update user profiles without a signed in user.", "displayName": "Read and write all users' full profiles", "id": "741f803b-c850-494e-b5df-cde7c675a1ca", "isEnabled": true, "origin": "Application", "value": "User.ReadWrite.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read user profiles without a signed in user.", "displayName": "Read all users' full profiles", "id": "df021288-bdef-4463-88db-98f22de89214", "isEnabled": true, "origin": "Application", "value": "User.Read.All" }, { "allowedMemberTypes": ["Application"], "description": "Allows the app to read and query your audit log activities, without a signed-in user.", "displayName": "Read all audit log data", "id": "b0afded3-3588-46d8-8b3d-9842eff778da", "isEnabled": true, "origin": "Application", "value": "AuditLog.Read.All" }], "oauth2PermissionScopes": [{ "adminConsentDescription": "Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.", "adminConsentDisplayName": "Maintain access to data you have given it access to", "id": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", "isEnabled": true, "type": "User", "userConsentDescription": "Allows the app to see and update the data you gave it access to, even when you are not currently using the app. This does not give the app any additional permissions.", "userConsentDisplayName": "Maintain access to data you have given it access to", "value": "offline_access" }, { "adminConsentDescription": "Allows the app to read the available Teams templates, on behalf of the signed-in user.", "adminConsentDisplayName": "Read available Teams templates", "id": "cd87405c-5792-4f15-92f7-debc0db6d1d6", "isEnabled": true, "type": "User", "userConsentDescription": "Read available Teams templates, on your behalf.", "userConsentDisplayName": "Read available Teams templates", "value": "TeamTemplates.Read" }] }];
  const applications: Application[] = [{ 'id': appId, 'requiredResourceAccess': [] }];
  const applicationPermissions = 'https://graph.microsoft.com/User.ReadWrite.All https://graph.microsoft.com/User.Read.All';
  const delegatedPermissions = 'https://graph.microsoft.com/offline_access';
  //#endregion

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(JSON.stringify({
      apps: [
        {
          appId: appId,
          name: 'CLI app1'
        }
      ]
    }));
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.patch,
      request.post,
      odata.getAllItems,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PERMISSION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds application permissions to appId without granting admin consent',
    async () => {
      jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
        switch (url) {
          case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
            return servicePrincipals;
          case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
            return applications;
          default:
            throw 'Invalid request';
        }
      });

      const patchStub = jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${applications[0].id}`) {
          return;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { applicationPermission: applicationPermissions, verbose: true } });
      assert(patchStub.called);
    }
  );

  it('adds application permissions to appId while granting admin consent',
    async () => {
      let amountOfPostCalls = 0;
      jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
        switch (url) {
          case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
            return servicePrincipals;
          case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
            return applications;
          default:
            throw 'Invalid request';
        }
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${applications[0].id}`) {
          return;
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/${servicePrincipalId}/appRoleAssignments`) {
          amountOfPostCalls += 1;
          return;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { applicationPermission: applicationPermissions, grantAdminConsent: true, verbose: true } });
      assert.strictEqual(amountOfPostCalls, 2);
    }
  );

  it('adds delegated permissions to appId without granting admin consent',
    async () => {
      jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
        switch (url) {
          case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
            return servicePrincipals;
          case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
            return applications;
          default:
            throw 'Invalid request';
        }
      });

      const patchStub = jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${applications[0].id}`) {
          return;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { delegatedPermission: delegatedPermissions, verbose: true } });
      assert(patchStub.called);
    }
  );

  it('adds delegated permissions to appId while granting admin consent',
    async () => {
      const pos: number = delegatedPermissions.lastIndexOf('/');
      const permissionName: string = delegatedPermissions.substring(pos + 1);

      jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
        switch (url) {
          case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
            return servicePrincipals;
          case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
            return applications;
          default:
            throw 'Invalid request';
        }
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${applications[0].id}`) {
          return;
        }
        throw 'Invalid request';
      });

      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants`) {
          return;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { delegatedPermission: delegatedPermissions, grantAdminConsent: true, verbose: true } });
      assert.deepStrictEqual(postStub.mock.lastCall[0].data, {
        clientId: servicePrincipalId,
        consentType: 'AllPrincipals',
        principalId: null,
        resourceId: 'fb4be1df-eaa6-4bd0-a068-71f9b2cbe2be',
        scope: permissionName
      });
    }
  );

  it('adds delegated and application permissions to appId while granting admin consent',
    async () => {
      let amountOfPostCalls = 0;

      jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
        switch (url) {
          case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
            return servicePrincipals;
          case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
            return applications;
          default:
            throw 'Invalid request';
        }
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/${applications[0].id}`) {
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/oauth2PermissionGrants`) {
          amountOfPostCalls++;
          return;
        }
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/servicePrincipals/${servicePrincipalId}/appRoleAssignments`) {
          amountOfPostCalls++;
          return;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { delegatedPermission: delegatedPermissions, applicationPermission: applicationPermissions, grantAdminConsent: true, verbose: true } });
      assert.strictEqual(amountOfPostCalls, 3);
    }
  );

  it('throws an error when application is not found', async () => {
    jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
          return [];
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { applicationPermission: applicationPermissions, verbose: true } }),
      new CommandError(`App with id ${appId} not found in Azure Active Directory`));
  });

  it('throws an error when service principal is not found', async () => {
    const api = 'https://grax.microsoft.com/User.ReadWrite.All';
    const pos: number = api.lastIndexOf('/');
    const servicePrincipalName: string = api.substring(0, pos);
    jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { applicationPermission: api, verbose: true } }),
      new CommandError(`Service principal ${servicePrincipalName} not found`));
  });

  it('throws an error when permission is not found', async () => {
    const api = 'https://graph.microsoft.com/NotFound.All';
    const pos: number = api.lastIndexOf('/');
    const servicePrincipalName: string = api.substring(0, pos);
    const permissionName: string = api.substring(pos + 1);
    jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url: string) => {
      switch (url) {
        case 'https://graph.microsoft.com/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames':
          return servicePrincipals;
        case `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '${appId}'&$select=id,requiredResourceAccess`:
          return applications;
        default:
          throw 'Invalid request';
      }
    });

    await assert.rejects(command.action(logger, { options: { applicationPermission: api, verbose: true } }),
      new CommandError(`Permission ${permissionName} for service principal ${servicePrincipalName} not found`));
  });

  it('passes validation if applicationPermission is passed', async () => {
    const actual = await command.validate({ options: { applicationPermission: applicationPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if delegatedPermission is passed', async () => {
    const actual = await command.validate({ options: { delegatedPermission: delegatedPermissions } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if both applicationPermission or delegatedPermission are passed',
    async () => {
      const actual = await command.validate({ options: { applicationPermission: applicationPermissions, delegatedPermission: delegatedPermissions } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation if both applicationPermission or delegatedPermission is not passed',
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
});