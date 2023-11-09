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
import command from './app-consent-set.js';

describe(commands.APP_CONSENT_SET, () => {
  //#region Mocked Responses
  const environmentName = 'Default-4be50206-9576-4237-8b17-38d8aadfaa36';
  const name = 'e0c89645-7f00-4877-a290-cbaf6e060da1';
  //#endregion

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
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
      request.post,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_CONSENT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: environmentName,
        name: 'invalid',
        bypass: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the name specified', async () => {
    const actual = await command.validate({
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before bypassing consent for the specified Microsoft Power App when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          environmentName: environmentName,
          name: name,
          bypass: true
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts bypassing the consent for the specified Microsoft Power App when confirm option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

      await command.action(logger, {
        options: {
          environmentName: environmentName,
          name: name,
          bypass: true
        }
      });
      assert(postSpy.notCalled);
    }
  );

  it('bypasses consent for the specified Microsoft Power App when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${environmentName}/apps/${name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`) {
          return { statusCode: 204 };
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await assert.doesNotReject(command.action(logger, {
        options: {
          debug: true,
          environmentName: environmentName,
          name: name,
          bypass: true
        }
      }));
    }
  );

  it('bypasses consent for the specified Microsoft Power App without prompting when confirm specified',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${environmentName}/apps/${name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`) {
          return { statusCode: 204 };
        }

        throw 'Invalid request';
      });

      await assert.doesNotReject(command.action(logger, {
        options: {
          environmentName: environmentName,
          name: name,
          bypass: true,
          force: true
        }
      }));
    }
  );

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: `Something went wrong bypassing the consent for the Microsoft Power App`
      }
    };

    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        environmentName: environmentName,
        name: name,
        bypass: true,
        force: true
      }
    } as any), new CommandError(error.error.message));
  });
});
