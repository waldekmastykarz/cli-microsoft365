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
import command from './managementapp-list.js';

describe(commands.MANAGEMENTAPP_LIST, () => {
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.put
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MANAGEMENTAPP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully retrieves management application', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return {
          "value": [{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }]);

    assert.strictEqual(actual, expected);
  });

  it('successfully retrieves multiple management applications', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return {
          "value": [{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }, { "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548f" }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }, { "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548f" }]);

    assert.strictEqual(actual, expected);
  });

  it('successfully handles no result found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return {
          "value": [{}]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([{}]);
    assert.strictEqual(actual, expected);
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
