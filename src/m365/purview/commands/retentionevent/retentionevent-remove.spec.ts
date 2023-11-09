import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './retentionevent-remove.js';

describe(commands.RETENTIONEVENT_REMOVE, () => {
  const validId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

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
    assert.strictEqual(command.name, commands.RETENTIONEVENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified retention event when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          id: validId
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified retention event when confirm option not passed and prompt not confirmed',
    async () => {
      const deleteSpy = jest.spyOn(request, 'delete').mockClear();
      await command.action(logger, {
        options: {
          id: validId
        }
      });
      assert(deleteSpy.notCalled);
    }
  );

  it('Correctly deletes retention event by id', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, {
      options: {
        id: validId
      }
    });
  });

  it('Correctly deletes retention event by id when prompt confirmed',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents/${validId}`) {
          return;
        }

        throw 'Invalid Request';
      });

      await command.action(logger, {
        options: {
          id: validId,
          force: true
        }
      });
    }
  );

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        code: "generalException",
        message: "Can't remove retention event"
      }
    };

    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        id: validId,
        force: true
      }
    }), new CommandError(error.error.message));
  });
});