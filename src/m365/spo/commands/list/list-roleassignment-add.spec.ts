import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import { RoleDefinition } from '../roledefinition/RoleDefinition.js';
import command from './list-roleassignment-add.js';

describe(commands.LIST_ROLEASSIGNMENT_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const userResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: 'i:0#.f|membership|john.doe@contoso.com',
    Title: 'John Doe',
    PrincipalType: 1,
    Email: 'john.doe@contoso.com',
    Expiration: '',
    IsEmailAuthenticationGuestUser: false,
    IsShareByEmailGuestUser: false,
    IsSiteAdmin: false,
    UserId: {
      NameId: '10032002473c5ae3',
      NameIdIssuer: 'urn:federation:microsoftonline'
    },
    UserPrincipalName: 'john.doe@contoso.com'
  };

  const groupResponse = {
    Id: 11,
    IsHiddenInUI: false,
    LoginName: "groupname",
    Title: "groupname",
    PrincipalType: 8,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    Description: "",
    OnlyAllowMembersViewMembership: true,
    OwnerTitle: "John Doe",
    RequestToJoinLeaveEmailSetting: null
  };

  const roledefinitionResponse: RoleDefinition = {
    BasePermissions: {
      High: 176,
      Low: 138612833
    },
    Description: "Can view pages and list items and download documents.",
    Hidden: false,
    Id: 1073741827,
    Name: "Read",
    Order: 128,
    RoleTypeKind: 2,
    BasePermissionsValue: [
      "ViewListItems",
      "OpenItems",
      "ViewVersions",
      "ViewFormPages",
      "Open",
      "ViewPages",
      "CreateSSCSite",
      "BrowseUserInfo",
      "UseClientIntegration",
      "UseRemoteAPIs",
      "CreateAlerts"
    ],
    RoleTypeKindValue: "Reader"
  };

  beforeAll(() => {
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      spo.getGroupByName,
      spo.getUserByEmail,
      spo.getRoleDefinitionByName
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_ROLEASSIGNMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 'abc', roleDefinitionId: 1073741827 } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the roleDefinitionId option is not a number',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 'abc' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the roleDefinitionId option is a number',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', principalId: 11, roleDefinitionId: 1073741827 } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('add role assignment on list by title and role definition id',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('_api/web/lists/getByTitle(\'test\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          listTitle: 'test',
          principalId: 11,
          roleDefinitionId: 1073741827
        }
      });
    }
  );

  it('add role assignment on list by id and role definition id', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment on list by url and role definition id', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetList(\'%2Fsites%2Fdocuments\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listUrl: 'sites/documents',
        principalId: 11,
        roleDefinitionId: 1073741827
      }
    });
  });

  it('add role assignment on list get principal id by upn', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    jest.spyOn(spo, 'getUserByEmail').mockClear().mockImplementation().resolves(userResponse);

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'User cannot be found.';
    jest.spyOn(spo, 'getUserByEmail').mockClear().mockImplementation().rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        upn: 'someaccount@tenant.onmicrosoft.com',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add role assignment on list get principal id by group name',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(spo, 'getGroupByName').mockClear().mockImplementation().resolves(groupResponse);

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
          groupName: 'someGroup',
          roleDefinitionId: 1073741827
        }
      });
    }
  );

  it('correctly handles error when group does not exist', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'Group cannot be found';
    jest.spyOn(spo, 'getGroupByName').mockClear().mockImplementation().rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        groupName: 'someGroup',
        roleDefinitionId: 1073741827
      }
    } as any), new CommandError(error));
  });

  it('add role assignment on list get role definition id by role definition name',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(spo, 'getRoleDefinitionByName').mockClear().mockImplementation().resolves(roledefinitionResponse);

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
          principalId: 11,
          roleDefinitionName: 'Full Control'
        }
      });
    }
  );

  it('correctly handles error when role definition does not exist',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web/lists(guid\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/roleassignments/addroleassignment(principalid=\'11\',roledefid=\'1073741827\')') > -1) {
          return;
        }

        throw 'Invalid request';
      });

      const error = 'No roledefinition is found for Read';
      jest.spyOn(spo, 'getRoleDefinitionByName').mockClear().mockImplementation().rejects(new Error(error));

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
          principalId: 11,
          roleDefinitionName: 'Full Control'
        }
      } as any), new CommandError(error));
    }
  );
});
