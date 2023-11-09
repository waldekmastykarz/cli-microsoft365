import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import ppSolutionGetCommand from './solution-get.js';
import command from './solution-remove.js';

describe(commands.SOLUTION_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '00000001-0000-0000-0001-00000000009b';
  const validName = 'Solution name';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  //#endregion

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let loggerLogToStderrSpy: jest.SpyInstance;

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
      powerPlatform.getDynamicsInstanceApiUrl,
      Cli.prompt,
      Cli.executeCommandWithOutput
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified solution owned by the currently signed-in user when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          environmentName: validEnvironment,
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

  it('aborts removing the specified solution owned by the currently signed-in user when confirm option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'delete').mockClear();

      await command.action(logger, {
        options: {
          environmentName: validEnvironment,
          id: validId
        }
      });
      assert(postSpy.notCalled);
    }
  );

  it('removes the specified solution owned by the currently signed-in user when prompt confirmed',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === ppSolutionGetCommand) {
          return ({
            stdout: `{
              "solutionid": "${validId}",
              "uniquename": "${validName}",
              "version": "1.0.0.0",
              "installedon": "2022-10-30T13:59:26Z",
              "solutionpackageversion": null,
              "friendlyname": "${validName}",
              "versionnumber": 1209676,
              "publisherid": {
                "friendlyname": "Default Publisher for org1547b730",
                "publisherid": "d21aab71-79e7-11dd-8874-00188b01e34f"
              }
            }`
          });
        }

        throw new CommandError('Unknown case');
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/solutions(${validId})`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));
      await command.action(logger, {
        options: {
          debug: true,
          environmentName: validEnvironment,
          name: validName
        }
      });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('removes the specified solution owned by the currently signed-in user without prompt for confirm',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/solutions(${validId})`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          environmentName: validEnvironment,
          id: validId,
          force: true
        }
      });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

    jest.spyOn(request, 'delete').mockClear().mockImplementation(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        id: validId,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
