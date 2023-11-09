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
import command from './user-get.js';

describe(commands.USER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    (command as any).items = [];
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
    assert.strictEqual(command.name.startsWith(commands.USER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email', 'job_title', 'state', 'url']);
  });

  it('calls user by e-mail', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/by_email.json?email=pl%40nubo.eu') {
        return Promise.resolve(
          [{ "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { email: "pl@nubo.eu" } } as any);
    assert.strictEqual(loggerLogSpy.mock.lastCall[0][0].id, 1496550646);
  });

  it('calls user by id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/1496550646.json') {
        return Promise.resolve(
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { id: 1496550646 } } as any);
    assert.strictEqual(loggerLogSpy.mock.lastCall[0].id, 1496550646);
  });

  it('calls the current user and json', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/current.json') {
        return Promise.resolve(
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.mock.lastCall[0].id, 1496550646);
  });

  it('correctly handles error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('correctly handles 404 error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      return Promise.reject({
        "statusCode": 404
      });
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Not found (404)'));
  });

  it('passes validation without parameters', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if id set ', async () => {
    const actual = await command.validate({ options: { id: 1496550646 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if email set', async () => {
    const actual = await command.validate({ options: { email: "pl@nubo.eu" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('does not pass with id and e-mail', async () => {
    const actual = await command.validate({ options: { id: 1496550646, email: "pl@nubo.eu" } }, commandInfo);
    assert.strictEqual(actual, "You are only allowed to search by ID or e-mail but not both");
  });
});
