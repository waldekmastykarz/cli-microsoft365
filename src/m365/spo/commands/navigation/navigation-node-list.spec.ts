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
import command from './navigation-node-list.js';

describe(commands.NAVIGATION_NODE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  const navigationNodeResponse = {
    value: [
      {
        "Id": 2003,
        "IsDocLib": true,
        "IsExternal": false,
        "IsVisible": true,
        "ListTemplateType": 0,
        "Title": "Node 1",
        "Url": "/sites/team-a/SitePages/page1.aspx"
      },
      {
        "Id": 2004,
        "IsDocLib": true,
        "IsExternal": false,
        "IsVisible": true,
        "ListTemplateType": 0,
        "Title": "Node 2",
        "Url": "/sites/team-a/SitePages/page2.aspx"
      }
    ]
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
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets nodes from the top navigation', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web/navigation/topnavigationbar') {
        return navigationNodeResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } });
    assert(loggerLogSpy.calledWith(navigationNodeResponse.value));
  });

  it('gets nodes from the quick launch', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web/navigation/quicklaunch') {
        return navigationNodeResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } });
    assert(loggerLogSpy.calledWith(navigationNodeResponse.value));
  });

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web/navigation/topnavigationbar') {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/web/navigation/topnavigationbar') {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when location is QuickLaunch and all required properties are present',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
