import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './roster-remove.js';

describe(commands.PLAN_REMOVE, () => {
  const validRosterId = 'CRp0hFSovEedkXtcX3WnS5gAGgch';

  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    promptOptions = undefined;
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the specified Roster when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          id: validRosterId
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified Roster when confirm option not passed and prompt not confirmed',
    async () => {
      const deleteSpy = jest.spyOn(request, 'delete').mockClear();
      await command.action(logger, {
        options: {
          id: validRosterId
        }
      });
      assert(deleteSpy.notCalled);
    }
  );

  it('correctly deletes Roster by id', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: validRosterId,
        force: true
      }
    });
  });

  it('correctly deletes Roster by id when prompt confirmed', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, {
      options: {
        id: validRosterId
      }
    });
  });

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects({
      error: {
        message: 'The requested item is not found.'
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: validRosterId,
        force: true
      }
    }), new CommandError('The requested item is not found.'));
  });
});
