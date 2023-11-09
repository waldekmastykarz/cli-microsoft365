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
import command from './roster-get.js';

describe(commands.ROSTER_GET, () => {
  const id = 'tYqYlNd6eECmsNhN_fcq85cAGAnd';
  const rosterGetResponse = {
    "id": id,
    "assignedSensitivityLabel": null
  };

  let log: any[];
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
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves Microsoft Planner Roster by specified id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${id}`) {
        return rosterGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: id, verbose: true } });
    assert(loggerLogSpy.calledWith(rosterGetResponse));
  });

  it('command correctly handles Microsoft Planner Roster get reject request',
    async () => {
      const errorMessage = 'Error: The requested item is not found.';
      jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${id}`) {
          throw errorMessage;
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          id: id
        }
      }), new CommandError(errorMessage));
    }
  );
});
