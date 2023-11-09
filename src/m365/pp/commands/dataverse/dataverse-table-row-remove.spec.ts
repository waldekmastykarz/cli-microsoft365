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
import command from './dataverse-table-row-remove.js';

describe(commands.DATAVERSE_TABLE_ROW_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validTableName = 'DataverseTable';
  const validEntitySetName = 'cr6c3_dataversetables';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const tableResponse = {
    EntitySetName: 'cr6c3_dataversetables'
  };
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
      request.get,
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
    assert.strictEqual(command.name, commands.DATAVERSE_TABLE_ROW_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID',
        tableName: validTableName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (tableName)',
    async () => {
      const actual = await command.validate({ options: { environmentName: validEnvironment, tableName: validTableName, id: validId } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if required options specified (entitySetName)',
    async () => {
      const actual = await command.validate({ options: { environmentName: validEnvironment, entitySetName: validEntitySetName, id: validId } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('prompts before removing the specified row from a dataverse table owned by the currently signed-in user when confirm option not passed',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

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

  it('aborts removing the specified row from a dataverse table owned by the currently signed-in user when confirm option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'delete').mockClear();
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: false }
      ));
      await command.action(logger, {
        options: {
          environmentName: validEnvironment,
          id: validId
        }
      });
      assert(postSpy.notCalled);
    }
  );

  it('removes the specified row according to the entitySetName parameter from a dataverse table owned by the currently signed-in user when prompt confirmed',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/${validEntitySetName}(${validId})`) {
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
          id: validId,
          entitySetName: validEntitySetName
        }
      });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('removes the specified row according to the tableName parameter from a dataverse table with the entitySetName parameter without confirmation prompt',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='${validTableName}')?$select=EntitySetName`)) {
          if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
            return tableResponse;
          }
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/${validEntitySetName}(${validId})`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          environmentName: validEnvironment,
          id: validId,
          tableName: validTableName,
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
        force: true,
        entitySetName: validEntitySetName
      }
    }), new CommandError(errorMessage));
  });

  it('removes dataverse table row with the entitySetName parameter and without confirmation',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/${validEntitySetName}(${validId})`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          environmentName: validEnvironment,
          id: validId,
          entitySetName: validEntitySetName,
          force: true
        }
      });

      assert(loggerLogToStderrSpy.called);
    }
  );
});