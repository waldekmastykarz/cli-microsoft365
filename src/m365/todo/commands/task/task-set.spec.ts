import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './task-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TASK_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let patchStub: jest.Mock<[options: CliRequestOptions]>;

  const getRequestData = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('4cb2b035-ad76-406c-bdc4-6c72ad403a22')/todo/lists",
    "value": [
      {
        "@odata.etag": "W/\"hHKQZHItDEOVCn8U3xuA2AABoWDAng==\"",
        "displayName": "Tasks List",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "none",
        "id": "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=="
      }
    ]
  };
  const patchRequestData = {
    "importance": "low",
    "isReminderOn": false,
    "status": "notStarted",
    "title": "New task",
    "createdDateTime": "2020-10-28T10:30:20.9783659Z",
    "lastModifiedDateTime": "2020-10-28T10:30:21.3616801Z",
    "id": "AAMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwBGAAAAAAAq-A2AAw08T7MU1EldWtTXBwCEcpBkci0MQ5UKfxTfG4DYAAGZB5U-AACEcpBkci0MQ5UKfxTfG4DYAAGhnfKPAAA=",
    "body": {
      "content": "",
      "contentType": "text"
    }
  };

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
    (command as any).items = [];
    patchStub = jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==/tasks/abc`) {
        return patchRequestData;
      }
      throw null;
    });


    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks%20List'`) {
        return getRequestData;
      }
      throw null;
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.patch,
      Date.now,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates tasks for  list using listId', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: "New task",
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(patchRequestData));
  });

  it('updates tasks for list using listName (debug)', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: "New task",
        listName: 'Tasks List',
        status: "notStarted",
        debug: true
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(patchRequestData));
  });

  it('updates tasks for list with bodyContent and bodyContentType',
    async () => {
      const bodyText = '<h3>Lorem ipsum</h3>';
      await command.action(logger, {
        options: {
          id: 'abc',
          title: 'New task',
          listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
          status: "notStarted",
          bodyContent: bodyText,
          bodyContentType: 'html'
        }
      } as any);

      assert.strictEqual(patchStub.mock.lastCall[0].data.body.content, bodyText);
      assert.strictEqual(patchStub.mock.lastCall[0].data.body.contentType, 'html');
    }
  );

  it('updates tasks for list with bodyContent and no bodyContentType',
    async () => {
      const bodyText = 'Lorem ipsum';
      await command.action(logger, {
        options: {
          id: 'abc',
          title: 'New task',
          listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
          status: "notStarted",
          bodyContent: bodyText
        }
      } as any);

      assert.strictEqual(patchStub.mock.lastCall[0].data.body.content, bodyText);
      assert.strictEqual(patchStub.mock.lastCall[0].data.body.contentType, 'text');
    }
  );

  it('updates tasks for list with importance', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        importance: 'high'
      }
    } as any);

    assert.strictEqual(patchStub.mock.lastCall[0].data.importance, 'high');
  });

  it('updates tasks for list with dueDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        dueDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(patchStub.mock.lastCall[0].data.dueDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('updates tasks for list with reminderDateTime', async () => {
    const dateTime = '2023-01-01T12:00:00';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        reminderDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(patchStub.mock.lastCall[0].data.reminderDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });


  it('updates To Do task with categories ', async () => {
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        status: "notStarted",
        categories: 'None,Preset24'
      }
    } as any);

    assert.deepStrictEqual(patchStub.mock.lastCall[0].data.categories, ['None', 'Preset24']);
  });

  it('updates To Do task with completedDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        completedDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(patchStub.mock.lastCall[0].data.completedDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('updates To Do task with startDateTime', async () => {
    const dateTime = '2023-01-01';
    await command.action(logger, {
      options: {
        id: 'abc',
        title: 'New task',
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
        startDateTime: dateTime
      }
    } as any);

    assert.deepStrictEqual(patchStub.mock.lastCall[0].data.startDateTime, { dateTime: dateTime, timeZone: 'Etc/GMT' });
  });

  it('rejects if no tasks list is found with the specified list name',
    async () => {
      jestUtil.restore(request.get);
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts: any) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks%20List'`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('4cb2b035-ad76-406c-bdc4-6c72ad403a22')/todo/lists",
            "value": []
          };
        }
        throw null;
      });

      await assert.rejects(command.action(logger, {
        options: {
          id: 'abc',
          title: "New task",
          listName: 'Tasks List',
          debug: true
        }
      } as any), new CommandError('The specified task list does not exist'));
    }
  );

  it('handles error correctly', async () => {
    jestUtil.restore(request.patch);
    jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        listName: "Tasks List",
        id: 'abc',
        title: "New task"
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if both listId and listName options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==',
          listName: 'Tasks List',
          title: 'New Task',
          id: 'abc'
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if neither listId nor listName options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          title: 'New Task',
          id: 'abc'
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if id not passed', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        title: 'New Task',
        listName: 'Tasks List'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if status is not allowed value', async () => {
    const options: any = {
      title: 'New Task',
      id: 'abc',
      status: "test",
      listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='

    };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, 'test is not a valid value. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred');
  });

  it('fails validation when invalid bodyContentType is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        id: 'abc',
        status: "notStarted",
        listName: 'Tasks List',
        bodyContentType: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid importance is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        id: 'abc',
        status: "notStarted",
        listName: 'Tasks List',
        importance: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid dueDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        id: 'abc',
        status: "notStarted",
        listName: 'Tasks List',
        dueDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid reminderDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        title: 'New task',
        id: 'abc',
        status: "notStarted",
        listName: 'Tasks List',
        reminderDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid completedDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        id: 'abc',
        title: 'New task',
        listName: 'Tasks List',
        completedDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid startDateTime is passed', async () => {
    const actual = await command.validate({
      options: {
        id: 'abc',
        title: 'New task',
        listName: 'Tasks List',
        startDateTime: '01/01/2022'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly validates the arguments', async () => {
    const actual = await command.validate({
      options: {
        title: 'New Task',
        id: 'abc',
        status: "notStarted",
        listId: 'AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=='
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
