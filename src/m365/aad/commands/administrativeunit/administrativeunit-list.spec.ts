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
import command from './administrativeunit-list.js';

describe(commands.ADMINISTRATIVEUNIT_LIST, () => {
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
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'visibility']);
  });

  it(`should get a list of administrative units`, async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits`) {
        return {
          value: [
            {
              id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
              displayName: 'European Division',
              visibility: 'HiddenMembership'
            },
            {
              id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
              displayName: 'Asian Division',
              visibility: null
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {}
    });

    assert(
      loggerLogSpy.calledWith([
        {
          id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
          displayName: 'European Division',
          visibility: 'HiddenMembership'
        },
        {
          id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
          displayName: 'Asian Division',
          visibility: null
        }
      ])
    );
  });

  it('handles error when retrieving administrative units list failed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits`) {
          throw { error: { message: 'An error has occurred' } };
        }
        throw `Invalid request`;
      });

      await assert.rejects(
        command.action(logger, { options: {} } as any),
        new CommandError('An error has occurred')
      );
    }
  );
});