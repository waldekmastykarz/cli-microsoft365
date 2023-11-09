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
import command from './list-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LIST_REMOVE, () => {
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
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.LIST_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a To Do task list by name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "FooList"
      }
    } as any);
    assert.strictEqual(log.length, 0);
  });

  it('removes a To Do task list by name when confirm option is passed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
            "value": [
              {
                "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
                "displayName": "FooList",
                "isOwner": true,
                "isShared": false,
                "wellknownListName": "none",
                "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
              }
            ]
          };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          name: "FooList",
          force: true
        }
      } as any);
      assert.strictEqual(log.length, 0);
    }
  );

  it('removes a To Do task list by id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
      }
    } as any);
    assert.strictEqual(log.length, 0);
  });

  it('handles error correctly when a list is not found for a specific name',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
            "value": []
          };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async () => {
        return;
      });
      await assert.rejects(command.action(logger, { options: { name: "FooList" } } as any), new CommandError('The list FooList cannot be found'));
    }
  );

  it('handles error correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { name: "FooList" } } as any), new CommandError('An error has occurred'));
  });

  it('prompts before removing the list when confirm option not passed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
            "value": [
              {
                "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
                "displayName": "FooList",
                "isOwner": true,
                "isShared": false,
                "wellknownListName": "none",
                "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
              }
            ]
          };
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: false }
      ));

      command.action(logger, {
        options: {
          name: "FooList"
        }
      } as any);
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }
      assert(promptIssued);
    }
  );

  it('fails validation if both name and id are not set', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        name: null,
        id: null
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all parameters are valid', async () => {
    const actual = await command.validate({
      options: {
        name: 'Foo'
      }
    }, commandInfo);

    assert.strictEqual(actual, true);
  });

  it('fails validation if both name and id are set', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        name: 'foo',
        id: 'bar'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
