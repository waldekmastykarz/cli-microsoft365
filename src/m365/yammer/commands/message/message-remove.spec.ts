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
import command from './message-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.MESSAGE_REMOVE, () => {
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { id: 'nonumber' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('id is required', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('calls the messaging endpoint with the right parameters and confirmation',
    async () => {
      const requestDeleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
          return;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, id: 10123190123123, force: true } });
      assert.strictEqual(requestDeleteStub.mock.lastCall[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
    }
  );

  it('calls the messaging endpoint with the right parameters without confirmation',
    async () => {
      const requestDeleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
          return;
        }
        throw 'Invalid request';
      });
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));

      await command.action(logger, { options: { debug: true, id: 10123190123123, force: false } });
      assert.strictEqual(requestDeleteStub.mock.lastCall[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
    }
  );

  it('does not call the messaging endpoint without confirmation', async () => {
    const requestDeleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });

    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: { debug: true, id: 10123190123123, force: false } });
    assert(requestDeleteStub.notCalled);
  });

  it('correctly handles error', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { id: 10123190123123, force: true } } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { id: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
