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
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './task-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TASK_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: true };
    });
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a To Do task by task id and task list name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('eded3a2a-8f01-40aa-998a-e4f02ec693ba')/todo/lists",
          "value": [
            {
              "@odata.etag": "W/\"tPAryi+qT0uvQKa/pHXU5AAAQchLxw==\"",
              "displayName": "Tasks",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "defaultList",
              "id": "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB="
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=/tasks/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=",
        listName: "Tasks"
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('removes a To Do task by task id and task list name when confirm option is passed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks'`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('eded3a2a-8f01-40aa-998a-e4f02ec693ba')/todo/lists",
            "value": [
              {
                "@odata.etag": "W/\"tPAryi+qT0uvQKa/pHXU5AAAQchLxw==\"",
                "displayName": "Tasks",
                "isOwner": true,
                "isShared": false,
                "wellknownListName": "defaultList",
                "id": "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB="
              }
            ]
          };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=/tasks/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=",
          listName: "Tasks",
          force: true
        }
      });
      assert.strictEqual(log.length, 0);
    }
  );

  it('removes a To Do task by task id and task list id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('eded3a2a-8f01-40aa-998a-e4f02ec693ba')/todo/lists/$entity",
          "@odata.etag": "W/\"tPAryi+qT0uvQKa/pHXU5AAAQchLxw==\"",
          "displayName": "Tasks",
          "isOwner": true,
          "isShared": false,
          "wellknownListName": "defaultList",
          "id": "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB="
        };
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=/tasks/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=",
        listId: "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB="
      }
    } as any);
    assert.strictEqual(log.length, 0);
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "error": {
            "code": "invalidRequest",
            "message": "The list FooList cannot be found",
            "innerError": {
              "code": "ErrorInvalidIdMalformed",
              "message": "Id is malformed.",
              "date": "2020-11-04T21:55:49",
              "request-id": "699fd167-e936-4d6c-8c4d-d616c758d7af",
              "client-request-id": "085a2508-e115-63b6-fcc9-05acc2133231"
            }
          }
        };
      }

      throw 'The list FooList cannot be found';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation(() => {
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=",
        listName: "FooList"
      }
    } as any), new CommandError('The list FooList cannot be found'));

  });

  it('prompts before removing the To Do task when confirm option not passed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'Tasks'`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('eded3a2a-8f01-40aa-998a-e4f02ec693ba')/todo/lists",
            "value": [
              {
                "@odata.etag": "W/\"tPAryi+qT0uvQKa/pHXU5AAAQchLxw==\"",
                "displayName": "Tasks",
                "isOwner": true,
                "isShared": false,
                "wellknownListName": "defaultList",
                "id": "BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB="
              }
            ]
          };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB=/tasks/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
        promptOptions = options;
        return { continue: false };
      });
      await command.action(logger, {
        options: {
          id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=",
          listName: "Tasks"
        }
      } as any);
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('passes validation when all parameters are valid with listId',
    async () => {
      const actual = await command.validate({
        options: {
          id: 'AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=',
          listId: 'BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB='
        }
      }, commandInfo);

      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when all parameters are valid with listName',
    async () => {
      const actual = await command.validate({
        options: {
          id: 'AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=',
          listName: 'Tasks'
        }
      }, commandInfo);

      assert.strictEqual(actual, true);
    }
  );

  it('fails validation if both listName and listId are not set', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=',
        listName: null,
        listId: null
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listName and listId are set', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=',
        listName: 'Tasks',
        listId: 'BBMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhBBB='
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
