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
import command from './task-checklistitem-remove.js';

describe(commands.TASK_CHECKLISTITEM_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validId = '71175';

  const responseChecklistWithId = {
    "71175": {
      "isChecked": false,
      "title": "test 2"
    }
  };
  const responseChecklistWithNoId = {
    "71176": {
      "isChecked": false,
      "title": "test 2"
    }
  };


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
    jest.spyOn(Cli.getInstance().config, 'all').mockClear().mockImplementation().value({});
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
      return { continue: true };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.patch,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_CHECKLISTITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removal when confirm option not passed', async () => {
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly deletes checklist item', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=checklist`) {
        return {
          "@odata.etag": "TestEtag",
          checklist: responseChecklistWithId
        };
      }
      throw 'Invalid Request';
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return;
      }
      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId,
        force: true
      }
    });
  });

  it('successfully remove checklist item with confirmation prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=checklist`) {
          return {
            "@odata.etag": "TestEtag",
            checklist: responseChecklistWithId
          };
        }
        throw 'Invalid Request';
      });
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
          return;
        }
        throw 'Invalid Request';
      });

      await command.action(logger, {
        options: {
          taskId: validTaskId,
          id: validId
        }
      });
    }
  );

  it('fails when checklist item does not exists', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=checklist`) {
        return {
          "@odata.etag": "TestEtag",
          checklist: responseChecklistWithNoId
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    }), new CommandError(`The specified checklist item with id ${validId} does not exist`));
  });

  it('correctly handles random API error', async () => {
    jestUtil.restore(request.get);
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
