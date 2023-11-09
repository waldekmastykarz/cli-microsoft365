import assert from 'assert';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import commands from './commands.js';
import command from './logout.js';

describe(commands.LOGOUT, () => {
  let log: string[];
  let logger: Logger;
  let authClearConnectionInfoStub: jest.Mock;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.resolve());
    authClearConnectionInfoStub = jest.spyOn(auth, 'clearConnectionInfo').mockClear().mockImplementation(() => Promise.resolve());
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
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

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGOUT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('logs out from Microsoft 365 when logged in', async () => {
    auth.service.connected = true;
    await command.action(logger, { options: { debug: true } });
    assert(!auth.service.connected);
  });

  it('logs out from Microsoft 365 when not logged in', async () => {
    auth.service.connected = false;
    await command.action(logger, { options: { debug: true } });
    assert(!auth.service.connected);
  });

  it('clears persisted connection info when logging out', async () => {
    auth.service.connected = true;
    await command.action(logger, { options: { debug: true } });
    assert(authClearConnectionInfoStub.called);
  });

  it('correctly handles error while clearing persisted connection info',
    async () => {
      jestUtil.restore(auth.clearConnectionInfo);
      jest.spyOn(auth, 'clearConnectionInfo').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));
      const logoutSpy = jest.spyOn(auth.service, 'logout').mockClear();
      auth.service.connected = true;

      try {
        await command.action(logger, { options: {} });
        assert(logoutSpy.called);
      }
      finally {
        jestUtil.restore([
          auth.clearConnectionInfo,
          auth.service.logout
        ]);
      }
    }
  );

  it('correctly handles error while clearing persisted connection info (debug)',
    async () => {
      jest.spyOn(auth, 'clearConnectionInfo').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));
      const logoutSpy = jest.spyOn(auth.service, 'logout').mockClear();
      auth.service.connected = true;

      try {
        await command.action(logger, { options: { debug: true } });
        assert(logoutSpy.called);
      }
      finally {
        jestUtil.restore([
          auth.clearConnectionInfo,
          auth.service.logout
        ]);
      }
    }
  );

  it('correctly handles error when restoring auth information', async () => {
    jestUtil.restore(auth.restoreAuth);
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));

    try {
      await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
    }
    finally {
      jestUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });
});
