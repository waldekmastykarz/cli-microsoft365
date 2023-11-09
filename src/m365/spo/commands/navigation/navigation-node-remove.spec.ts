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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './navigation-node-remove.js';

describe(commands.NAVIGATION_NODE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let promptOptions: any;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(spo, 'getRequestDigest').mockClear().mockImplementation().resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      request.post,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.NAVIGATION_NODE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes navigation node from the top navigation', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar/getbyid(2003)`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('removes navigation node from the top navigation (debug)', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar/getbyid(2003)`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', force: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('prompts before removing navigation node when confirmation argument not passed',
    async () => {
      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing app when prompt not confirmed', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(() => {
      throw 'Invalid request';
    });
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } });
  });

  it('removes the navigation node when prompt confirmed', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/navigation/topnavigationbar/getbyid(2003)`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(() => {
      throw { error: 'An error has occurred' };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', location: 'TopNavigationBar', id: '2003' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified location is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'invalid', id: '2003' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when location is TopNavigationBar and all required properties are present',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'TopNavigationBar', id: '2003' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when location is QuickLaunch and all required properties are present',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', location: 'QuickLaunch', id: '2003' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
