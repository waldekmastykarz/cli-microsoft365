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
import command from './user-remove.js';

describe(commands.USER_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validUsername = 'john.doe@contoso.com';
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptOptions: any;

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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName is not a valid upn', async () => {
    const actual = await command.validate({
      options: {
        userName: 'Invalid upn'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (userId)', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (userName)',
    async () => {
      const actual = await command.validate({ options: { userName: validUsername } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('prompts before removing the specified user when confirm option not passed',
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

  it('aborts removing the specified user when confirm option not passed and prompt not confirmed',
    async () => {
      const deleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation().resolves();

      await command.action(logger, {
        options: {
          id: validId
        }
      });
      assert(deleteStub.notCalled);
    }
  );

  it('removes the specified user by id when prompt confirmed', async () => {
    const deleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${validId}`) {
        return;
      }

      throw 'Invalid request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, {
      options: {
        verbose: true,
        id: validId
      }
    });
    assert(deleteStub.called);
  });

  it('removes the specified user by userName when prompt confirmed',
    async () => {
      const deleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users/${validUsername}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          verbose: true,
          userName: validUsername
        }
      });
      assert(deleteStub.called);
    }
  );

  it('removes the specified user by Username without confirmation prompt',
    async () => {
      const deleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users/${validUsername}`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          verbose: true,
          userName: validUsername,
          force: true
        }
      });
      assert(deleteStub.called);
    }
  );

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: 'The user cannot be found.'
      }
    };

    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        id: validId,
        force: true
      }
    }), new CommandError(error.error.message));
  });
});
