import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './m365group-recyclebinitem-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.M365GROUP_RECYCLEBINITEM_REMOVE, () => {
  let cli: Cli;
  const validGroupId = '00000000-0000-0000-0000-000000000000';
  const validGroupDisplayName = 'Dev Team';
  const validGroupMailNickname = 'Devteam';

  const singleGroupsResponse = {
    value: [
      {
        id: validGroupId,
        displayName: validGroupDisplayName,
        mailNickname: validGroupDisplayName,
        mail: 'Devteam@contoso.com',
        groupTypes: [
          "Unified"
        ]
      }
    ]
  };

  const multipleGroupsResponse = {
    value: [
      {
        id: validGroupId,
        displayName: validGroupDisplayName,
        mailNickname: validGroupDisplayName,
        mail: 'Devteam@contoso.com',
        groupTypes: [
          "Unified"
        ]
      },
      {
        id: validGroupId,
        displayName: validGroupDisplayName,
        mailNickname: validGroupDisplayName,
        mail: 'Devteam@contoso.com',
        groupTypes: [
          "Unified"
        ]
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  beforeAll(() => {
    cli = Cli.getInstance();
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
    promptOptions = undefined;
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      Cli.prompt,
      Cli.handleMultipleResultsFound,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id is not a valid GUID', async () => {
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
        id: validGroupId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified group when confirm option not passed with id',
    async () => {
      await command.action(logger, {
        options: {
          id: validGroupId
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified group when confirm option not passed and prompt not confirmed',
    async () => {
      const deleteSpy = jest.spyOn(request, 'delete').mockClear();
      await command.action(logger, {
        options: {
          id: validGroupId
        }
      });
      assert(deleteSpy.notCalled);
    }
  );

  it('throws error message when no group was found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname,
        force: true
      }
    }), new CommandError(`The specified group '${validGroupMailNickname}' does not exist.`));
  });

  it('throws error message when multiple groups were found', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return multipleGroupsResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname,
        force: true
      }
    }), new CommandError("Multiple groups with name 'Devteam' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt',
    async () => {
      let removeRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupDisplayName)}'`) {
          return multipleGroupsResponse;
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
          removeRequestIssued = true;
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves(singleGroupsResponse.value[0]);

      jestUtil.restore(Cli.prompt);

      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          displayName: validGroupDisplayName
        }
      });
      assert(removeRequestIssued);
    }
  );

  it('correctly deletes group by id with confirm flag', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: validGroupId,
        force: true
      }
    });
  });

  it('correctly deletes group by id', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, {
      options: {
        id: validGroupId
      }
    });
  });

  it('correctly deletes group by displayName', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupDisplayName)}'`) {
        return singleGroupsResponse;
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, {
      options: {
        displayName: validGroupDisplayName
      }
    });
  });

  it('correctly deletes group by mailNickname', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'`) {
        return singleGroupsResponse;
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validGroupId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        mailNickname: validGroupMailNickname
      }
    });
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'An error has occurred'
      }
    };
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        id: validGroupId,
        force: true
      }
    }), new CommandError("An error has occurred"));
  });
}); 
