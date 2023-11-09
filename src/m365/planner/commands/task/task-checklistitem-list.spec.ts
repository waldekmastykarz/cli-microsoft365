import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './task-checklistitem-list.js';

describe(commands.TASK_CHECKLISTITEM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  const jsonOutput = {
    "checklist": {
      "33224": {
        "isChecked": false,
        "title": "Some checklist",
        "orderHint": "8585576049720396756P(",
        "lastModifiedDateTime": "2022-02-04T19:12:53.4692149Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
          }
        }
      },
      "69115": {
        "isChecked": false,
        "title": "Some checklist more",
        "orderHint": "85855760494@",
        "lastModifiedDateTime": "2022-02-04T19:12:55.4735671Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
          }
        }
      }
    }
  };
  const textOutput = {
    "checklist": [{
      "id": "33224",
      "isChecked": false,
      "title": "Some checklist",
      "orderHint": "8585576049720396756P(",
      "lastModifiedDateTime": "2022-02-04T19:12:53.4692149Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
        }
      }
    },
    {
      "id": "69115",
      "isChecked": false,
      "title": "Some checklist more",
      "orderHint": "85855760494@",
      "lastModifiedDateTime": "2022-02-04T19:12:55.4735671Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
        }
      }
    }
    ]
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_CHECKLISTITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'isChecked']);
  });

  it('successfully handles item found(JSON)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details?$select=checklist`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true
      }
    });
    assert(loggerLogSpy.calledWith(jsonOutput.checklist));
  });

  it('successfully handles item found(TEXT)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details?$select=checklist`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true, output: 'text'
      }
    });
    assert(loggerLogSpy.calledWith(textOutput.checklist));
  });

  it('correctly handles item not found', async () => {
    jestUtil.restore(request.get);
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('The requested item is not found.'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('The requested item is not found.'));
  });

  it('correctly handles random API error', async () => {
    jestUtil.restore(request.get);
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
