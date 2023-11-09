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
import command from './app-list.js';

describe(commands.APP_LIST, () => {
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
    assert.strictEqual(command.name, commands.APP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['appId', 'id', 'displayName', 'signInAudience']);
  });

  it(`should get a list of Azure AD app registrations`, async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications`) {
        return {
          value: [
            {
              id: '340a4aa3-1af6-43ac-87d8-189819003952',
              appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
              displayName: 'My App 1',
              description: 'My second app',
              signInAudience: 'My Audience'
            },
            {
              id: '340a4aa3-1af6-43ac-87d8-189819003953',
              appId: '9b1b1e42-794b-4c71-93ac-5ed92488b670',
              displayName: 'My App 2',
              description: 'My second app',
              signInAudience: 'My Audience'
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
          id: '340a4aa3-1af6-43ac-87d8-189819003952',
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          displayName: 'My App 1',
          description: 'My second app',
          signInAudience: 'My Audience'
        },
        {
          id: '340a4aa3-1af6-43ac-87d8-189819003953',
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b670',
          displayName: 'My App 2',
          description: 'My second app',
          signInAudience: 'My Audience'
        }
      ])
    );
  });

  it('handles error when retrieving app list failed', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred')
    );
  });
});
