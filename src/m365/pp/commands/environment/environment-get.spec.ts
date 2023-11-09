import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './environment-get.js';

describe(commands.ENVIRONMENT_GET, () => {
  const environmentName = 'Default-de347bc8-1aeb-4406-8cb3-97db021cadb4';
  const environmentResponse = {
    "id": `/providers/Microsoft.BusinessAppPlatform/environments/Default-de347bc8-1aeb-4406-8cb3-97db021cadb4`,
    "type": "Microsoft.BusinessAppPlatform/environments",
    "location": "unitedstates",
    "name": "Default-de347bc8-1aeb-4406-8cb3-97db021cadb4",
    "properties": {
      "displayName": "contoso (default)",
      "isDefault": true
    }
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
      pid.getProcessName,
      session.getId
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENVIRONMENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'id']);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = `Resource '' does not exist or one of its queried reference-property objects are not present`;
    jest.spyOn(request, 'get').mockClear().mockImplementation(async () => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: environmentName
      }
    }), new CommandError(errorMessage));
  });

  it('retrieves Microsoft Power Platform environment by name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(environmentName)}?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return environmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: environmentName,
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(environmentResponse));
  });

  it('retrieves default Microsoft Power Platform environment', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/~Default?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return environmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(environmentResponse));
  });

  it('retrieves Microsoft Power Platform environment as Admin', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${formatting.encodeQueryParameter(environmentName)}?api-version=2020-10-01`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return environmentResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: environmentName,
        asAdmin: true,
        verbose: true
      }
    });

    assert(loggerLogSpy.calledWith(environmentResponse));
  });
});
