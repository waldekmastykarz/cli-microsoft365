import assert from 'assert';
import auth, { AuthType, CloudType } from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { accessToken } from '../../utils/accessToken.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import commands from './commands.js';
import command from './status.js';

describe(commands.STATUS, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let loggerLogToStderrSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.resolve());
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      auth.ensureAccessToken,
      accessToken.getUserNameFromAccessToken
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.STATUS), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows logged out status when not logged in', async () => {
    auth.service.connected = false;
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith('Logged out'));
  });

  it('shows logged out status when not logged in (verbose)', async () => {
    auth.service.connected = false;
    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogToStderrSpy.calledWith('Logged out from Microsoft 365'));
  });

  it('shows logged out status when the refresh token is expired', async () => {
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: 'abc',
      accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc'
    };

    auth.service.connected = true;
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: {} }), new CommandError(`Your login has expired. Sign in again to continue. Error`));
  });

  it('shows logged out status when refresh token is expired (debug)',
    async () => {
      auth.service.accessTokens['https://graph.microsoft.com'] = {
        expiresOn: 'abc',
        accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc'
      };

      auth.service.connected = true;
      const error = new Error('Error');
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => { return Promise.reject(error); });
      await assert.rejects(command.action(logger, { options: { debug: true } }), new CommandError(`Your login has expired. Sign in again to continue. Error`));
      assert(loggerLogToStderrSpy.calledWith(error));
    }
  );

  it('shows logged in status when logged in', async () => {
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: 'abc',
      accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ing0NTh4eU9wbHNNMkg3TlhrMlN4MTd4MXVwYyIsImtpZCI6Ing0NTh4eU9wbHNNMkg3TlhrN1N4MTd4MXVwYyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLndpbmRvd3MubmV0IiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvY2FlZTMyZTYtNDA1ZC00MjRhLTljZjEtMjA3MWQwNDdmMjk4LyIsImlhdCI6MTUxNTAwNDc4NCwibmJmIjoxNTE1MDA0Nzg0LCJleHAiOjE1MTUwMDg2ODQsImFjciI6IjEiLCJhaW8iOiJBQVdIMi84R0FBQUFPN3c0TDBXaHZLZ1kvTXAxTGJMWFdhd2NpOEpXUUpITmpKUGNiT2RBM1BvPSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiIwNGIwNzc5NS04ZGRiLTQ2MWEtYmJlZS0wMmY5ZTFiZjdiNDYiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaXBhZGRyIjoiOC44LjguOCIsIm5hbWUiOiJKb2huIERvZSIsIm9pZCI6ImYzZTU5NDkxLWZjMWEtNDdjYy1hMWYwLTk1ZWQ0NTk4MzcxNyIsInB1aWQiOiIxMDk0N0ZGRUE2OEJDQ0NFIiwic2NwIjoiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwic3ViIjoiemZicmtUV1VQdEdWUUg1aGZRckpvVGp3TTBrUDRsY3NnLTJqeUFJb0JuOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImNhZWUzM2U2LTQwNWQtNDU0YS05Y2YxLTMwNzFkMjQxYTI5OCIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AY29udG9zby5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFUZVdpelVmUTBheFBLMVRUVXhsQUEiLCJ2ZXIiOiIxLjAifQ==.abc'
    };

    auth.service.connected = true;
    auth.service.authType = AuthType.DeviceCode;
    auth.service.appId = '8dd76117-ab8e-472c-b5c1-a50e13b457cd';
    auth.service.tenant = 'common';
    auth.service.cloudType = CloudType.Public;
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
    jest.spyOn(accessToken, 'getUserNameFromAccessToken').mockClear().mockImplementation(() => { return 'admin@contoso.onmicrosoft.com'; });
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      connectedAs: 'admin@contoso.onmicrosoft.com',
      authType: 'DeviceCode',
      appId: '8dd76117-ab8e-472c-b5c1-a50e13b457cd',
      appTenant: 'common',
      cloudType: 'Public'
    }));
  });

  it('correctly reports access token', async () => {
    auth.service.connected = true;
    auth.service.authType = AuthType.DeviceCode;
    auth.service.appId = '8dd76117-ab8e-472c-b5c1-a50e13b457cd';
    auth.service.tenant = 'common';
    auth.service.cloudType = CloudType.Public;
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
    jest.spyOn(accessToken, 'getUserNameFromAccessToken').mockClear().mockImplementation(() => { return 'admin@contoso.onmicrosoft.com'; });
    auth.service.accessTokens = {
      'https://graph.microsoft.com': {
        expiresOn: '123',
        accessToken: 'abc'
      }
    };
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogToStderrSpy.calledWith({
      connectedAs: 'admin@contoso.onmicrosoft.com',
      authType: 'DeviceCode',
      appId: '8dd76117-ab8e-472c-b5c1-a50e13b457cd',
      appTenant: 'common',
      accessTokens: '{\n  "https://graph.microsoft.com": {\n    "expiresOn": "123",\n    "accessToken": "abc"\n  }\n}',
      cloudType: 'Public'
    }));
  });

  it('correctly reports access token - no user', async () => {
    auth.service.connected = true;
    auth.service.authType = AuthType.DeviceCode;
    auth.service.appId = '8dd76117-ab8e-472c-b5c1-a50e13b457cd';
    auth.service.tenant = 'common';
    auth.service.cloudType = CloudType.Public;
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
    auth.service.accessTokens = {
      'https://graph.microsoft.com': {
        expiresOn: '123',
        accessToken: 'abc'
      }
    };

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogToStderrSpy.calledWith({
      connectedAs: '',
      authType: 'DeviceCode',
      appId: '8dd76117-ab8e-472c-b5c1-a50e13b457cd',
      appTenant: 'common',
      accessTokens: '{\n  "https://graph.microsoft.com": {\n    "expiresOn": "123",\n    "accessToken": "abc"\n  }\n}',
      cloudType: 'Public'
    }));
  });

  it('correctly handles error when restoring auth', async () => {
    jestUtil.restore(auth.restoreAuth);
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
