import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
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
import command from './externalconnection-get.js';

describe(commands.EXTERNALCONNECTION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  const externalConnection: ExternalConnectors.ExternalConnection =
  {
    "id": "contosohr",
    "name": "Contoso HR",
    "description": "Connection to index Contoso HR system",
    "state": "draft",
    "configuration": {
      "authorizedAppIds": [
        "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
      ]
    }
  };

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
    assert.strictEqual(command.name, commands.EXTERNALCONNECTION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
      }
    }), new CommandError('An error has occurred'));
  });

  it('should get external connection information for the Microsoft Search by id (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
          return externalConnection;
        }
        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          id: 'contosohr'
        }
      });

      const call: sinon.SinonSpyCall = loggerLogSpy.mock.lastCall;
      assert.strictEqual(call.mock.calls[0].id, 'contosohr');
      assert.strictEqual(call.mock.calls[0].name, 'Contoso HR');
      assert.strictEqual(call.mock.calls[0].description, 'Connection to index Contoso HR system');
      assert.strictEqual(call.mock.calls[0].state, 'draft');
    }
  );

  it('should get external connection information for the Microsoft Search by name',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
          return {
            "value": [
              externalConnection
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          name: 'Contoso HR'
        }
      });
      const call: sinon.SinonSpyCall = loggerLogSpy.mock.lastCall;
      assert.strictEqual(call.mock.calls[0].id, 'contosohr');
      assert.strictEqual(call.mock.calls[0].name, 'Contoso HR');
      assert.strictEqual(call.mock.calls[0].description, 'Connection to index Contoso HR system');
      assert.strictEqual(call.mock.calls[0].state, 'draft');
    }
  );

  it('fails retrieving external connection not found by name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return {
          "value": []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso HR'
      }
    }), new CommandError(`External connection with name 'Contoso HR' not found`));
  });
});
