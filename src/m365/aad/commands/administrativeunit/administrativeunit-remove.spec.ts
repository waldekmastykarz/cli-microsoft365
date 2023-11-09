import assert from 'assert';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './administrativeunit-remove.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.ADMINISTRATIVEUNIT_REMOVE, () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const secondAdministrativeUnitId = 'fc33aa61-cf0e-1234-9506-f633347202ab';
  const displayName = 'European Division';
  const invalidDisplayName = 'European';

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
      request.delete,
      request.get,
      Cli.handleMultipleResultsFound,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified administrative unit by id without prompting for confirmation',
    async () => {
      const deleteRequestStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { id: administrativeUnitId, force: true } });
      assert(deleteRequestStub.called);
    }
  );

  it('removes the specified administrative unit by displayName while prompting for confirmation',
    async () => {
      const getRequestStub = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
          return {
            value: [
              { id: administrativeUnitId }
            ]
          };
        }

        throw 'Invalid Request';
      });

      const deleteRequestStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { displayName: displayName } });
      assert(deleteRequestStub.called);
      assert(getRequestStub.called);
    }
  );

  it('removes selected administrative unit when more administrative units with the specified displayName found while prompting for confirmation',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
          return {
            value: [
              {
                id: administrativeUnitId
              },
              {
                id: secondAdministrativeUnitId
              }
            ]
          };
        }

        throw 'Invalid Request';
      });

      const deleteRequestStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });
      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({ id: administrativeUnitId });

      await command.action(logger, { options: { displayName: displayName } });
      assert(deleteRequestStub.called);
    }
  );

  it('throws an error when administrative unit by id cannot be found',
    async () => {
      const error = {
        error: {
          code: 'Request_ResourceNotFound',
          message: `Resource '${administrativeUnitId}' does not exist or one of its queried reference-property objects are not present.`,
          innerError: {
            date: '2023-10-27T12:24:36',
            'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
            'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
          }
        }
      };
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${administrativeUnitId}`) {
          throw error;
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { id: administrativeUnitId, force: true } }),
        new CommandError(error.error.message));
    }
  );

  it('prompts before removing the specified administrative unit when confirm option not passed',
    async () => {
      await command.action(logger, { options: { id: administrativeUnitId } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing administrative unit when prompt not confirmed',
    async () => {
      const deleteSpy = jest.spyOn(request, 'delete').mockClear().mockImplementation().resolves();

      await command.action(logger, { options: { id: administrativeUnitId } });
      assert(deleteSpy.notCalled);
    }
  );

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: administrativeUnitId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error message when no administrative unit was found by displayName',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'&$select=id`) {
          return { value: [] };
        }

        throw 'Invalid Request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await assert.rejects(command.action(logger, { options: { displayName: invalidDisplayName } }), new CommandError(`The specified administrative unit '${invalidDisplayName}' does not exist.`));
    }
  );
});