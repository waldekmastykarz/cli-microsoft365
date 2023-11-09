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
import command from './group-list.js';

describe(commands.GROUP_LIST, () => {
  const groupsResponse = [{
    Id: 15,
    Title: "Contoso Members",
    LoginName: "Contoso Members",
    "Description": "SharePoint Contoso",
    IsHiddenInUI: false,
    PrincipalType: 8
  }];
  const groupsResponseValue = {
    value: groupsResponse
  };
  const associatedGroupsResponse = {
    "AssociatedMemberGroup":
    {
      "Id": 6,
      "Title": "Site Members",
      "LoginName": "Site Members",
      "Description": "",
      "IsHiddenInUI": false,
      "PrincipalType": 8
    },
    "AssociatedOwnerGroup": {
      "Id": 7,
      "Title": "Site Owners",
      "LoginName": "Site Owners",
      "Description": "",
      "IsHiddenInUI": false,
      "PrincipalType": 8
    },
    "AssociatedVisitorGroup": {
      "Id": 8,
      "Title": "Site Visitors",
      "LoginName": "Site Visitors",
      "Description": "",
      "IsHiddenInUI": false,
      "PrincipalType": 8
    }
  };
  const associatedGroupsResponseText = [{
    "Id": 6,
    "Title": "Site Members",
    "LoginName": "Site Members",
    "Description": "",
    "IsHiddenInUI": false,
    "PrincipalType": 8,
    "Type": "AssociatedMemberGroup"
  },
  {
    "Id": 7,
    "Title": "Site Owners",
    "LoginName": "Site Owners",
    "Description": "",
    "IsHiddenInUI": false,
    "PrincipalType": 8,
    "Type": "AssociatedOwnerGroup"
  },
  {
    "Id": 8,
    "Title": "Site Visitors",
    "LoginName": "Site Visitors",
    "Description": "",
    "IsHiddenInUI": false,
    "PrincipalType": 8,
    "Type": "AssociatedVisitorGroup"
  }
  ];



  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'LoginName', 'IsHiddenInUI', 'PrincipalType', 'Type']);
  });

  it('retrieves all site groups', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        return groupsResponseValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledOnceWith(groupsResponse));
  });

  it('retrieves associated groups from the site', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web?$expand=') > -1) {
        return JSON.stringify(associatedGroupsResponse);
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroupsOnly: true
      }
    });
    assert(loggerLogSpy.calledOnceWith(JSON.stringify(associatedGroupsResponse)));
  });

  it('retrieves associated groups from the site with return type json',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web?$expand=') > -1) {
          return JSON.stringify(associatedGroupsResponse);
        }
        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          webUrl: 'https://contoso.sharepoint.com',
          associatedGroupsOnly: true,
          output: 'json'
        }
      });
      assert(loggerLogSpy.calledOnceWith(JSON.stringify(associatedGroupsResponse)));
    }
  );

  it('retrieves associated groups from the site with return type text',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web?$expand=') > -1) {
          return associatedGroupsResponse;
        }
        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          webUrl: 'https://contoso.sharepoint.com',
          associatedGroupsOnly: true,
          output: 'text'
        }
      });
      assert(loggerLogSpy.calledOnceWith(associatedGroupsResponseText));
    }
  );

  it('command correctly handles group list reject request', async () => {
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
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation url is valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
