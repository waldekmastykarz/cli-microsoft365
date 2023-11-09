import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './list-get.js';

describe(commands.LIST_GET, () => {
  let commandInfo: CommandInfo;
  const validName: string = "Task list";
  const validId: string = "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=";
  const listResponse = {
    value: [
      {
        displayName: "test cli",
        isOwner: true,
        isShared: false,
        wellknownListName: "none",
        id: "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA="
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

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
    assert.strictEqual(command.name, commands.LIST_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'id']);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws an error when no list found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(validName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ value: [] });
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validName
      }
    }), new CommandError(`The specified list '${validName}' does not exist.`));
  });

  it('lists a specific To Do task list based on the id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/${validId}`) {
        return (listResponse.value[0]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: validId
      }
    });

    assert(loggerLogSpy.calledWith(listResponse.value[0]));
  });

  it('lists a specific To Do task list based on the name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(validName)}'`) {
        return (listResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: validName
      }
    });

    assert(loggerLogSpy.calledWith(listResponse.value[0]));
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async () => {
      throw { error: { message: 'An error has occurred' } };
    });

    await assert.rejects(command.action(logger, { options: { id: validId } } as any), new CommandError('An error has occurred'));
  });
});
