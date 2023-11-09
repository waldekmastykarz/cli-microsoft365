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
import command from './group-member-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.GROUP_MEMBER_ADD, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  const jsonSingleUser =
  {
    ErrorMessage: null,
    IconUrl: "https://contoso.sharepoint.com/sites/SiteA/_layouts/15/images/siteicon.png",
    InvitedUsers: null,
    Name: "Site A",
    PermissionsPageRelativeUrl: null,
    StatusCode: 0,
    UniquelyPermissionedUsers: [],
    Url: "https://contoso.sharepoint.com/sites/SiteA",
    UsersAddedToGroup: [
      {
        AllowedRoles: [
          0
        ],
        CurrentRole: 0,
        DisplayName: "Alex Wilber",
        Email: "Alex.Wilber@contoso.com",
        InvitationLink: null,
        IsUserKnown: true,
        Message: null,
        Status: true,
        User: "i:0#.f|membership|Alex.Wilber@contoso.com"
      }
    ]
  };

  const jsonGenericError =
  {
    ErrorMessage: "The selected permission level is not valid.",
    IconUrl: null,
    InvitedUsers: null,
    Name: null,
    PermissionsPageRelativeUrl: null,
    StatusCode: -63,
    UniquelyPermissionedUsers: null,
    Url: null,
    UsersAddedToGroup: null
  };

  const groupResponse = {
    Id: 32
  };

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
      request.post,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both groupId and groupName options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          groupName: "Contoso Site Owners",
          userNames: "Alex.Wilber@contoso.com"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both groupId and groupName options are not passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          userNames: "Alex.Wilber@contoso.com"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both userNames and emails options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          emails: "Alex.Wilber@contoso.com",
          userNames: "Alex.Wilber@contoso.com"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both userNames and userIds options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          userIds: 5,
          userNames: "Alex.Wilber@contoso.com"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both emails and aadGroupIds options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          emails: "Alex.Wilber@contoso.com",
          aadGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both userIds and aadGroupNames options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          userIds: 5,
          aadGroupNames: "Azure AD Group name"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both userIds and emails options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          userIds: 5,
          emails: "Alex.Wilber@contoso.com"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if userNames, emails, userIds, aadGroupIds or aadGroupNames options are not passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if webURL is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "InvalidWEBURL", groupId: 32, userNames: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if groupID is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: "NOGROUP", userNames: "Alex.Wilber@contoso.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userIds is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userIds: "9,invalidUserId" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userNames is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userNames: "Alex.Wilber@contoso.com,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if emails is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, emails: "Alex.Wilber@contoso.com,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if aadGroupIds is Invalid', async () => {
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, aadGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983,9" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all the required options are specified',
    async () => {
      const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, userNames: "Alex.Wilber@contoso.com" } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['DisplayName', 'Email']);
  });

  it('adds user to a SharePoint Group by groupId and userNames', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return jsonSingleUser;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userNames: "Alex.Wilber@contoso.com"
      }
    });
    assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
  });

  it('adds user to a SharePoint Group by groupId and userIds (Debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
          opts.data) {
          return jsonSingleUser;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/siteusers/GetById('9')?$select=AadObjectId`) {
          return {
            AadObjectId: {
              NameId: '6cc1797e-5463-45ec-bb1a-b93ec198bab6',
              NameIdIssuer: 'urn:federation:microsoftonline'
            }
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
          return groupResponse;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      await command.action(logger, {
        options: {
          debug: true,
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          userIds: 9
        }
      });
      assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
    }
  );

  it('adds user to a SharePoint Group by groupId and userNames (Debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
          opts.data) {
          return jsonSingleUser;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
          return groupResponse;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          userNames: "Alex.Wilber@contoso.com"
        }
      });
      assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
    }
  );

  it('adds user to a SharePoint Group by groupName and emails (DEBUG)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq 'Alex.Wilber%40contoso.com'&$select=id`) {
          return { value: [{ id: "2056d2f6-3257-4253-8cfc-b73393e414e5" }] };
        }

        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetByName(`) > -1) {
          return {
            Id: 7
          };
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
          opts.data) {
          return jsonSingleUser;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      await command.action(logger, {
        options: {
          debug: true,
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupName: "Contoso Site Owners",
          emails: "Alex.Wilber@contoso.com"
        }
      });
      assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
    }
  );

  it('adds user to a SharePoint Group by groupId and aadGroupIds (Debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
          opts.data) {
          return jsonSingleUser;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
          return groupResponse;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          aadGroupIds: "56ca9023-3449-4e98-a96a-69e81a6f4983"
        }
      });
      assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
    }
  );

  it('adds user to a SharePoint Group by groupId and aadGroupNames (Debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
          opts.data) {
          return jsonSingleUser;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Azure%20AD%20Group%20name'&$select=id`) {
          return {
            value: [{
              id: 'Group name'
            }]
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
          return groupResponse;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      await command.action(logger, {
        options: {
          debug: true,
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          aadGroupNames: "Azure AD Group name"
        }
      });
      assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
    }
  );

  it('fails to get group when does not exists', async () => {
    const errorMessage = 'The specified group does not exist in the SharePoint site';
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetByName('`) > -1) {
        throw { error: { 'odata.error': { message: { value: errorMessage } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupName: "Contoso Site Owners",
        emails: "Alex.Wilber@contoso.com"
      }
    }), new CommandError(errorMessage));
  });

  it('handles generic error when adding user to a SharePoint Group by groupId and userIds',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/siteusers/GetById('9')`) {
          throw 'User not found';
        }

        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
          return groupResponse;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          userIds: 9
        }
      }), new CommandError(`Resource '9' does not exist.`));
    }
  );

  it('handles error when adding user to SharePoint Group group', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
        opts.data) {
        return jsonGenericError;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/sites/SiteA",
        groupId: 32,
        userNames: 'Alex.Wilber@contoso.com'
      }
    }), new CommandError('The selected permission level is not valid.'));
  });

  it('handles error when multiple groups with the specified displayName found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Azure%20AD%20Group%20name'&$select=id`) {
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
          webUrl: "https://contoso.sharepoint.com/sites/SiteA",
          groupId: 32,
          aadGroupNames: "Azure AD Group name"
        }
      }), new CommandError("Resource 'Azure AD Group name' does not exist."));
    }
  );

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Azure%20AD%20Group%20name'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/sites/SiteA/_api/web/sitegroups/GetById('32')?$select=Id`) {
          return groupResponse;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({ id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });

      jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://contoso.sharepoint.com/sites/SiteA/_api/SP.Web.ShareObject' &&
          opts.data) {
          return jsonSingleUser;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await command.action(logger, {
        options: {
          debug: true, webUrl: "https://contoso.sharepoint.com/sites/SiteA", groupId: 32, aadGroupNames: "Azure AD Group name"
        }
      });
      assert(loggerLogSpy.calledWith(jsonSingleUser.UsersAddedToGroup));
    }
  );
});
