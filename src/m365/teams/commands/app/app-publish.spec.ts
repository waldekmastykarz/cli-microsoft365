import assert from 'assert';
import fs from 'fs';
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
import command from './app-publish.js';

describe(commands.APP_PUBLISH, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  const appResponse = {
    id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
    externalId: "b5561ec9-8cab-4aa3-8aa2-d8d7172e4311",
    displayName: "Test App",
    distributionMethod: "organization"
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      fs.readFileSync,
      fs.existsSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PUBLISH);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the filePath does not exist', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
    const actual = await command.validate({
      options: { filePath: 'invalid.zip' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the filePath points to a directory', async () => {
    const stats: fs.Stats = new fs.Stats();
    jest.spyOn(stats, 'isDirectory').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'lstatSync').mockClear().mockReturnValue(stats);

    const actual = await command.validate({
      options: { filePath: './' }
    }, commandInfo);
    jestUtil.restore([
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const stats: fs.Stats = new fs.Stats();
    jest.spyOn(stats, 'isDirectory').mockClear().mockReturnValue(false);
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'lstatSync').mockClear().mockReturnValue(stats);

    const actual = await command.validate({
      options: {
        filePath: 'teamsapp.zip'
      }
    }, commandInfo);
    jestUtil.restore([
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('adds new Teams app to the tenant app catalog', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return appResponse;
      }

      throw 'Invalid request';
    });

    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('123');

    await command.action(logger, { options: { filePath: 'teamsapp.zip' } });
    assert(loggerLogSpy.calledWith(appResponse));
  });

  it('adds new Teams app to the tenant app catalog (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return appResponse;
      }

      throw 'Invalid request';
    });

    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('123');

    await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip' } });
    assert(loggerLogSpy.calledWith(appResponse));
  });

  it('correctly handles error when publishing an app', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    });


    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('123');

    await assert.rejects(command.action(logger, { options: { filePath: 'teamsapp.zip' } } as any), new CommandError('An error has occurred'));
  });
});
