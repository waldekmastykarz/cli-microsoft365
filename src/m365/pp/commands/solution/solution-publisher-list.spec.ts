import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './solution-publisher-list.js';

describe(commands.SOLUTION_PUBLISHER_LIST, () => {
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const publisherResponse = {
    "value": [
      {
        "publisherid": "00000001-0000-0000-0000-00000000005a",
        "uniquename": "Cree38e",
        "friendlyname": "CDS Default Publisher",
        "versionnumber": 1074060,
        "isreadonly": false,
        "description": null,
        "customizationprefix": "cr6c3",
        "customizationoptionvalueprefix": 43186
      },
      {
        "publisherid": "d21aab70-79e7-11dd-8874-00188b01e34f",
        "uniquename": "MicrosoftCorporation",
        "friendlyname": "MicrosoftCorporation",
        "versionnumber": 1226559,
        "isreadonly": false,
        "customizationprefix": "",
        "customizationoptionvalueprefix": 0
      }
    ]
  };

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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_PUBLISHER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['publisherid', 'uniquename', 'friendlyname']);
  });

  it('retrieves publishers from power platform environment', async () => {
    jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          return publisherResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment } });
    assert(loggerLogSpy.calledWith(publisherResponse.value));
  });

  it('retrieves publishers from power platform environment including the Microsoft Publishers',
    async () => {
      jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`)) {
          if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
            return publisherResponse;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, environmentName: validEnvironment, includeMicrosoftPublishers: true } });
      assert(loggerLogSpy.calledWith(publisherResponse.value));
    }
  );

  it('correctly handles API OData error', async () => {
    jest.spyOn(powerPlatform, 'getDynamicsInstanceApiUrl').mockClear().mockImplementation(async () => envUrl);

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'&api-version=9.1`)) {
        if ((opts.headers?.accept as string).indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, InvalidOperationException',
                message: {
                  value: `Resource '' does not exist or one of its queried reference-property objects are not present`
                }
              }
            }
          };
        }
      }

    });

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironment } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });
});
