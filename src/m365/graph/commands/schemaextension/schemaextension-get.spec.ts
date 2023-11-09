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
import command from './schemaextension-get.js';

describe(commands.SCHEMAEXTENSION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

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
  });


  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });
  it('gets schema extension', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        id: 'adatumisv_exo2'
      }
    });
    try {
      assert(loggerLogSpy.calledWith({
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
        "id": "adatumisv_exo2",
        "description": "sample description",
        "targetTypes": [
          "Message"
        ],
        "status": "Available",
        "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
        "properties": [
          {
            "name": "p1",
            "type": "String"
          },
          {
            "name": "p2",
            "type": "String"
          }
        ]
      }));
    }
    finally {
      jestUtil.restore(request.get);
    }
  });
  it('gets schema extension(debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    });
    assert(loggerLogSpy.calledWith({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
      "id": "adatumisv_exo2",
      "description": "sample description",
      "targetTypes": [
        "Message"
      ],
      "status": "Available",
      "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
      "properties": [
        {
          "name": "p1",
          "type": "String"
        },
        {
          "name": "p2",
          "type": "String"
        }
      ]
    }));
  });
  it('handles error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
