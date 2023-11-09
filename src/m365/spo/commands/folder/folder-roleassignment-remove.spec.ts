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
import spoGroupGetCommand from '../group/group-get.js';
import spoUserGetCommand from '../user/user-get.js';
import command from './folder-roleassignment-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FOLDER_ROLEASSIGNMENT_REMOVE, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptOptions: any;

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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      Cli.executeCommandWithOutput,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ROLEASSIGNMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation if the principalId option is not a number',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 'abc' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if principalId and upn are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, upn: 'someaccount@tenant.onmicrosoft.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if principalId and groupName are specified',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', principalId: 11, groupName: 'someGroup' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if upn and groupName are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission', upn: 'someaccount@tenant.onmicrosoft.com', groupName: 'someGroup' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither upn nor principalId or groupName is specified',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents/FolderPermission' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if folderUrl is not specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', upn: 'someaccount@tenant.onmicrosoft.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('remove role assignment from folder by folderUrl', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        principalId: 11,
        force: true
      }
    });
  });

  it('remove role assignment from folder and get principal id by upn',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoUserGetCommand) {
          return {
            stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
          };
        }

        throw new CommandError('Unknown case');
      });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          folderUrl: '/Shared Documents/FolderPermission',
          upn: 'someaccount@tenant.onmicrosoft.com',
          force: true
        }
      });
    }
  );

  it('correctly handles error when upn does not exist', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no user found';
    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoUserGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        upn: 'someaccount@tenant.onmicrosoft.com',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('remove role assignment from folder and get principal id by group name',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoGroupGetCommand) {
          return {
            stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
          };
        }

        throw new CommandError('Unknown case');
      });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          folderUrl: '/Shared Documents/FolderPermission',
          groupName: 'someGroup',
          force: true
        }
      });
    }
  );

  it('correctly handles error when group does not exist', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    const error = 'no group found';
    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
        throw error;
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup',
        force: true
      }
    } as any), new CommandError(error));
  });

  it('aborts removing role assignment when prompt not confirmed', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      }
    });

    assert(requests.length === 0);
  });

  it('prompts before removing role assignment when confirmation argument not passed',
    async () => {
      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com',
          folderUrl: '/Shared Documents/FolderPermission',
          groupName: 'someGroup'
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }
      assert(promptIssued);
    }
  );

  it('removes role assignment when prompt confirmed', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2FShared%20Documents%2FFolderPermission\')/ListItemAllFields/roleassignments/removeroleassignment(principalid=\'11\')') {
        return;
      }

      throw 'Invalid request';
    });

    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoGroupGetCommand) {
        return {
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents/FolderPermission',
        groupName: 'someGroup'
      }
    });
  });
});
