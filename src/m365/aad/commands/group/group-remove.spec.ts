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
import { aadGroup } from '../../../../utils/aadGroup.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './group-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.GROUP_REMOVE, () => {
  const groupId = '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a';
  const displayName = 'CLI Test Group';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  let cli: Cli;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    cli = Cli.getInstance();
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
      request.get,
      request.delete,
      aadGroup.getGroupIdByDisplayName,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified group by id without prompting for confirmation',
    async () => {
      const deleteRequestStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, id: groupId, force: true } });
      assert(deleteRequestStub.called);
    }
  );

  it('removes the specified group by displayName while prompting for confirmation',
    async () => {
      jest.spyOn(aadGroup, 'getGroupIdByDisplayName').mockClear().mockImplementation().resolves(groupId);

      const deleteRequestStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { verbose: true, displayName: displayName } });
      assert(deleteRequestStub.called);
    }
  );

  it('throws an error when group by id cannot be found', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `Resource '${groupId}' does not exist or one of its queried reference-property objects are not present.`,
        innerError: {
          date: '2023-08-30T14:32:41',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };

    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${groupId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, id: groupId, force: true } }),
      new CommandError(error.error.message));
  });

  it('prompts before removing the specified group when confirm option not passed',
    async () => {
      await command.action(logger, { options: { id: groupId } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('handles error when multiple groups with the specified displayName found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        return 'Invalid Request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects('DELETE request executed');

      await assert.rejects(command.action(logger, {
        options: {
          displayName: displayName,
          force: true
        }
      }), new CommandError(`Multiple groups with name 'CLI Test Group' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
    }
  );

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({ id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });

      const deleteRequestStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups/9b1b1e42-794b-4c71-93ac-5ed92488b67f`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { displayName: displayName, force: true } });
      assert(deleteRequestStub.called);
    }
  );

  it('aborts removing group when prompt not confirmed', async () => {
    const deleteSpy = jest.spyOn(request, 'delete').mockClear().mockImplementation().resolves();

    await command.action(logger, { options: { id: groupId } });
    assert(deleteSpy.notCalled);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
