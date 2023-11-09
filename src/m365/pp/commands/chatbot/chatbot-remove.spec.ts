import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import ppChatbotGetCommand from './chatbot-get.js';
import command from './chatbot-remove.js';

describe(commands.CHATBOT_REMOVE, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validName = 'CLI 365 Chatbot';
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
      request.post,
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
    assert.strictEqual(command.name, commands.CHATBOT_REMOVE);
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

  it('prompts before removing the specified chatbot owned by the currently signed-in user when confirm option not passed',
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

  it('aborts removing the specified chatbot owned by the currently signed-in user when confirm option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();

      await command.action(logger, {
        options: {
          environmentName: validEnvironment,
          id: validId
        }
      });
      assert(postSpy.notCalled);
    }
  );

  it('removes the specified chatbot owned by the currently signed-in user when prompt confirmed by name',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === ppChatbotGetCommand) {
          return ({
            stdout: `{ "authenticationtrigger": 0, "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f", "statuscode": 1, "createdon": "2022-11-19T10:42:22Z", "statecode": 0, "schemaname": "new_bot_23f5f58697fd43d595eb451c9797a53d", "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "name": "CLI 365 Chatbot", "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238", "ismanaged": false, "versionnumber": 1429641, "publishedon": "2022-11-19T19:19:53Z", "timezoneruleversionnumber": 0, "language": 1033, "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "overwritetime": "1900-01-01T00:00:00Z", "modifiedon": "2022-11-19T20:19:57Z", "componentstate": 0, "botid": "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893", "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "componentidunique": "cdcd6496-e25d-4ad1-91cf-3f4d547fdd23", "authenticationmode": 1, "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f", "accesscontrolpolicy": 0, "runtimeprovider": 0, "_publishedby_value": null, "authenticationconfiguration": null, "authorizedsecuritygroupids": null, "overriddencreatedon": null, "applicationmanifestinformation": null, "importsequencenumber": null, "synchronizationstatus": null, "_modifiedonbehalfby_value": null, "template": null, "_providerconnectionreferenceid_value": null, "configuration": null, "utcconversiontimezonecode": null, "_createdonbehalfby_value": null, "iconbase64": null, "supportedlanguages": null, "_owningteam_value": null, "iscustomizable": { "Value": true, "CanBeChanged": true, "ManagedPropertyLogicalName": "iscustomizableanddeletable" } }`
          });
        }

        throw new CommandError('Unknown case');
      });

      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots(${validId})/Microsoft.Dynamics.CRM.PvaDeleteBot?tag=deprovisionbotondelete`) {
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
          verbose: true,
          environmentName: validEnvironment,
          name: validName
        }
      });
      assert(postStub.called);
    }
  );

  it('removes the specified chatbot without confirmation prompt by id',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots(${validId})/Microsoft.Dynamics.CRM.PvaDeleteBot?tag=deprovisionbotondelete`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          verbose: true,
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

    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { throw { error: { error: { message: errorMessage } } }; });

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
