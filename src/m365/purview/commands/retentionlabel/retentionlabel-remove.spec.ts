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
import command from './retentionlabel-remove.js';

describe(commands.RETENTIONLABEL_REMOVE, () => {
  const validId = 'e554d69c-0992-4f9b-8a66-fca3c4d9c531';

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
    assert.strictEqual(command.name, commands.RETENTIONLABEL_REMOVE);
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

  it('prompts before removing the specified retention label when confirm option not passed',
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

  it('aborts removing the specified retention label when confirm option not passed and prompt not confirmed',
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

  it('Correctly deletes retention label by id', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        id: validId
      }
    });
  });

  it('Correctly deletes retention label by id when prompt confirmed',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels/${validId}`) {
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
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        id: validId,
        force: true
      }
    }), new CommandError("An error has occurred"));
  });
});