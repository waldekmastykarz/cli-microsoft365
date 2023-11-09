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
import command from './listitem-record-unlock.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LISTITEM_RECORD_UNLOCK, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const listUrl = "/MyLibrary";
  const listTitle = "MyLibrary";
  const listId = "cc27a922-8224-4296-90a5-ebbc54da2e85";
  const webUrl = "https://contoso.sharepoint.com";
  const listResponse = {
    "RootFolder": {
      "Exists": true,
      "IsWOPIEnabled": false,
      "ItemCount": 0,
      "Name": listTitle,
      "ProgID": null,
      "ServerRelativeUrl": listUrl,
      "TimeCreated": "2019-01-11T10:03:19Z",
      "TimeLastModified": "2019-01-11T10:03:20Z",
      "UniqueId": listId,
      "WelcomePage": ""
    }
  };

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.resolve());
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
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
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_RECORD_UNLOCK);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('unlocks a list item based on listUrl (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        verbose: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('unlocks a list item based on listTitle', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('${listTitle}')/?$expand=RootFolder&$select=RootFolder`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('unlocks a list item based on listId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'${listId}')/?$expand=RootFolder&$select=RootFolder`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        listId: listId,
        webUrl: webUrl,
        listItemId: 1
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1
      }
    }), new CommandError(errorMessage));
  });

  it('fails validation if both id and title options are not passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 1 } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', listItemId: 1, listTitle: listTitle } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 1 } }, commandInfo);
      assert(actual);
    }
  );

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: '12345', listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 1 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both id and title options are passed', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not passed', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 'abc', listTitle: listTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});