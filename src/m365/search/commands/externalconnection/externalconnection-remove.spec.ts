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
import command from './externalconnection-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.EXTERNALCONNECTION_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      cli.getSettingWithDefaultValue,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.EXTERNALCONNECTION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the specified external connection by id when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          id: "contosohr"
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing the specified external connection by name when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          name: "Contoso HR"
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified external connection when confirm option not passed and prompt not confirmed (debug)',
    async () => {
      const postSpy = jest.spyOn(request, 'delete').mockClear();
      await command.action(logger, { options: { debug: true, id: "contosohr" } });
      assert(postSpy.notCalled);
    }
  );

  it('removes the specified external connection when prompt confirmed (debug)',
    async () => {
      let externalConnectionRemoveCallIssued = false;

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
          externalConnectionRemoveCallIssued = true;
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));


      await command.action(logger, { options: { debug: true, id: "contosohr" } });
      assert(externalConnectionRemoveCallIssued);
    }
  );

  it('removes the specified external connection without prompting when confirm specified',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { id: "contosohr", force: true } });
    }
  );

  it('removes external connection with specified ID', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return;
      }
      throw '';
    });

    await command.action(logger, { options: { id: "contosohr", force: true } });
  });

  it('removes external connection with specified name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts: any) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq `) > -1) {
        return {
          value: [
            {
              "id": "contosohr",
              "name": "Contoso HR",
              "description": "Connection to index Contoso HR system"
            }
          ]
        };
      }
      throw '';
    });

    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
        return;
      }
      throw '';
    });

    await command.action(logger, { options: { name: "Contoso HR", force: true } });
  });

  it('fails to get external connection by name when it does not exists',
    async () => {
      jestUtil.restore(request.get);
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts: any) => {
        if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
          return { value: [] };
        }

        throw 'The specified connection does not exist in Microsoft Search';
      });

      await assert.rejects(command.action(logger, {
        options: {
          name: "Fabrikam HR",
          force: true
        }
      } as any), new CommandError("The specified connection does not exist in Microsoft Search"));
    }
  );

  it('fails when multiple external connections with same name exists',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jestUtil.restore(request.get);
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=`) > -1) {
          return {
            value: [
              {
                "id": "fabrikamhr"
              },
              {
                "id": "contosohr"
              }
            ]
          };
        }

        throw "Invalid request";
      });

      await assert.rejects(command.action(logger, {
        options: {
          name: "My HR",
          force: true
        }
      } as any), new CommandError("Multiple external connections with name My HR found. Found: fabrikamhr, contosohr."));
    }
  );

  it('handles selecting single result when external connections with the specified name found and cli is set to prompt',
    async () => {
      let removeRequestIssued = false;

      jestUtil.restore(request.get);
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/external/connections?$filter=name eq 'My%20HR'&$select=id`) {
          return {
            value: [
              {
                "id": "fabrikamhr"
              },
              {
                "id": "contosohr"
              }
            ]
          };
        }

        throw "Invalid request";
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({
        "id": "contosohr"
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts: any) => {
        if (opts.url === 'https://graph.microsoft.com/v1.0/external/connections/contosohr') {
          removeRequestIssued = true;
          return;
        }
        throw '';
      });

      await command.action(logger, { options: { name: "My HR", force: true } });
      assert(removeRequestIssued);
    }
  );
});
