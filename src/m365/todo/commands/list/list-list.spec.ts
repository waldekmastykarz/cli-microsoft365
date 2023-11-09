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
import command from './list-list.js';

describe(commands.LIST_LIST, () => {
  let log: string[];
  let logger: Logger;

  beforeAll(() => {
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'id']);
  });

  it('lists To Do task lists', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpQ==\"",
              "displayName": "Tasks",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "defaultList",
              "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpw==\"",
              "displayName": "Foo",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrqQ==\"",
              "displayName": "Bar",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([
      {
        "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpQ==\"",
        "displayName": "Tasks",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "defaultList",
        "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
      },
      {
        "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpw==\"",
        "displayName": "Foo",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "none",
        "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
      },
      {
        "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrqQ==\"",
        "displayName": "Bar",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "none",
        "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
      }
    ]);
    assert.strictEqual(actual, expected);
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
