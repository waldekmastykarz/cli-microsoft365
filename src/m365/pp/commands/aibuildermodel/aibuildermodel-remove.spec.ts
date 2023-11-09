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
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import ppAiBuilderModelGetCommand from './aibuildermodel-get.js';
import command from './aibuildermodel-remove.js';

describe(commands.AIBUILDERMODEL_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '08ffffbe-ec1c-4e64-b64b-dd1db926c613';
  const validName = 'CLI 365 Ai Builder Model';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const aiBuilderModelResponse = `{
    "statecode": 0,
    "_msdyn_templateid_value": "10707e4e-1d56-e911-8194-000d3a6cd5a5",
    "msdyn_modelcreationcontext": "{}",
    "createdon": "2022-11-29T11:58:45Z",
    "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "modifiedon": "2022-11-29T11:58:45Z",
    "msdyn_sharewithorganizationoncreate": false,
    "msdyn_aimodelidunique": "b0328b67-47e2-4202-8189-e617ec9a88bd",
    "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238",
    "ismanaged": false,
    "versionnumber": 1458121,
    "msdyn_name": "Document Processing 11/29/2022, 12:58:43 PM",
    "introducedversion": "1.0",
    "statuscode": 0,
    "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "overwritetime": "1900-01-01T00:00:00Z",
    "componentstate": 0,
    "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
    "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
    "msdyn_aimodelid": "08ffffbe-ec1c-4e64-b64b-dd1db926c613",
    "_msdyn_activerunconfigurationid_value": null,
    "overriddencreatedon": null,
    "_msdyn_retrainworkflowid_value": null,
    "importsequencenumber": null,
    "_msdyn_scheduleinferenceworkflowid_value": null,
    "_modifiedonbehalfby_value": null,
    "utcconversiontimezonecode": null,
    "_createdonbehalfby_value": null,
    "_owningteam_value": null,
    "timezoneruleversionnumber": null,
    "iscustomizable": {
      "Value": true,
      "CanBeChanged": true,
      "ManagedPropertyLogicalName": "iscustomizableanddeletable"
    }
  }`;
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
    assert.strictEqual(command.name, commands.AIBUILDERMODEL_REMOVE);
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

  it('prompts before removing the specified AI builder model owned by the currently signed-in user when confirm option not passed',
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

  it('aborts removing the specified AI builder model owned by the currently signed-in user when confirm option not passed and prompt not confirmed',
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

  it('removes the specified AI builder model owned by the currently signed-in user when prompt confirmed',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === ppAiBuilderModelGetCommand) {
          return ({
            stdout: aiBuilderModelResponse
          });
        }

        throw new CommandError('Unknown case');
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels(${validId})`) {
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

  it('removes the specified AI builder model without confirmation prompt',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/msdyn_aimodels(${validId})`) {
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
