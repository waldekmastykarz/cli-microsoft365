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
import command from './externalconnection-list.js';

describe(commands.EXTERNALCONNECTION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  const externalConnections: { value: ExternalConnectors.ExternalConnection[] } = {
    value: [
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
      }
    ]
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
    assert.strictEqual(command.name, commands.EXTERNALCONNECTION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'state']);
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

  it('retrieves list of external connections defined in the Microsoft Search',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts: any) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
          return externalConnections;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true } } as any);
      assert(loggerLogSpy.calledWith(externalConnections.value));
    }
  );
});
