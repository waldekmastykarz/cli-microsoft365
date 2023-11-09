import assert from 'assert';
import fs from 'fs';
import auth, { AuthType, CloudType } from '../../Auth.js';
import { CommandArgs, CommandError } from '../../Command.js';
import { Cli } from '../../cli/Cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import commands from './commands.js';
import command from './login.js';

describe(commands.LOGIN, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.resolve());
    jest.spyOn(auth, 'clearConnectionInfo').mockClear().mockImplementation(() => Promise.resolve());
    jest.spyOn(auth, 'storeConnectionInfo').mockClear().mockImplementation(() => Promise.resolve());
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
    commandInfo = Cli.getCommandInfo(command);
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
    jest.spyOn(auth.service, 'logout').mockClear().mockImplementation(() => { });
  });

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      auth.service.logout,
      auth.ensureAccessToken
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGIN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('in telemetry defaults to Public cloud when no cloud has been specified',
    () => {
      const args: CommandArgs = { options: {} };
      command.telemetry.forEach(fn => fn(args));
      assert.strictEqual((command as any).telemetryProperties.cloud, CloudType.Public);
    }
  );

  it('in telemetry tracks the specified cloud', () => {
    const args: CommandArgs = { options: { cloud: 'USGov' } };
    command.telemetry.forEach(fn => fn(args));
    assert.strictEqual((command as any).telemetryProperties.cloud, 'USGov');
  });

  it('logs in to Microsoft 365', async () => {
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
    await command.action(logger, { options: {} });
    assert(auth.service.connected);
  });

  it('logs in to Microsoft 365 (debug)', async () => {
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
    await command.action(logger, { options: { debug: true } });
    assert(auth.service.connected);
  });

  it('logs in to Microsoft 365 using username and password when authType password set',
    async () => {
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
      await command.action(logger, { options: { authType: 'password', userName: 'user', password: 'password' } });
      assert.strictEqual(auth.service.authType, AuthType.Password, 'Incorrect authType set');
      assert.strictEqual(auth.service.userName, 'user', 'Incorrect user name set');
      assert.strictEqual(auth.service.password, 'password', 'Incorrect password set');
    }
  );

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is provided',
    async () => {
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => 'certificate');

      await assert.rejects(command.action(logger, { options: { authType: 'certificate', certificateFile: 'certificate' } }));
      assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
      assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
    }
  );

  it('logs in to Microsoft 365 using certificate when authType certificate set with thumbprint',
    async () => {
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => 'certificate');

      await assert.rejects(command.action(logger, { options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } }));
      assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
      assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
      assert.strictEqual(auth.service.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
    }
  );

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is provided',
    async () => {
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => 'certificate');

      await assert.rejects(command.action(logger, { options: { authType: 'certificate', certificateBase64Encoded: 'certificate', thumbprint: 'thumbprint' } }));
      assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
      assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
      assert.strictEqual(auth.service.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
    }
  );

  it('logs in to Microsoft 365 using system managed identity when authType identity set',
    async () => {
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));

      await command.action(logger, { options: { authType: 'identity', userName: 'ac9fbed5-804c-4362-a369-21a4ec51109e' } });
      assert.strictEqual(auth.service.authType, AuthType.Identity, 'Incorrect authType set');
      assert.strictEqual(auth.service.userName, 'ac9fbed5-804c-4362-a369-21a4ec51109e', 'Incorrect userName set');
    }
  );

  it('logs in to Microsoft 365 using user-assigned managed identity when authType identity set',
    async () => {
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));

      await command.action(logger, { options: { authType: 'identity' } });
      assert.strictEqual(auth.service.authType, AuthType.Identity, 'Incorrect authType set');
      assert.strictEqual(auth.service.userName, undefined, 'Incorrect userName set');
    }
  );


  it('logs in to Microsoft 365 using client secret authType "secret" set',
    async () => {
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
      await command.action(logger, { options: { authType: 'secret', secret: 'unBrEakaBle@123' } });
      assert.strictEqual(auth.service.authType, AuthType.Secret, 'Incorrect authType set');
      assert.strictEqual(auth.service.secret, 'unBrEakaBle@123', 'Incorrect secret set');
    }
  );

  it('logs in to Microsoft 365 using the specified cloud', async () => {
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));
    await command.action(logger, { options: { cloud: 'USGov' } });
    assert.strictEqual(auth.service.cloudType, CloudType.USGov);
  });

  it('supports specifying authType', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--authType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying userName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--userName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying password', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--password') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if invalid authType specified', async () => {
    const actual = await command.validate({ options: { authType: 'invalid authType' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and userName and password not specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'password' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if authType is set to password and userName not specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'password', password: 'password' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if authType is set to password and password not specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'password', userName: 'user' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if authType is set to certificate and both certificateFile and certificateBase64Encoded are specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate', certificateBase64Encoded: 'certificateB64', thumbprint: 'thumbprint' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if authType is set to certificate and neither certificateFile nor certificateBase64Encoded are specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'certificate' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if authType is set to certificate and certificateFile does not exist',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation cloud is set to an invalid value', async () => {
    const actual = await command.validate({ options: { cloud: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if authType is set to certificate and certificateFile and thumbprint are specified',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if authType is set to certificate and certificateFile are specified',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if authType is set to password and userName and password specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'password', userName: 'user', password: 'password' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if authType is set to deviceCode and userName and password not specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'deviceCode' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if authType is not set and userName and password not specified',
    async () => {
      const actual = await command.validate({ options: {} }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when cloud is set to Public', async () => {
    const actual = await command.validate({ options: { cloud: 'Public' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to USGov', async () => {
    const actual = await command.validate({ options: { cloud: 'USGov' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to USGovHigh', async () => {
    const actual = await command.validate({ options: { cloud: 'USGovHigh' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to USGovDoD', async () => {
    const actual = await command.validate({ options: { cloud: 'USGovDoD' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to China', async () => {
    const actual = await command.validate({ options: { cloud: 'China' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles error in device code auth flow', async () => {
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Error'));
  });

  it('correctly handles error in device code auth flow (debug)', async () => {
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('Error'));
  });

  it('logs in to Microsoft 365 using browser authentication', async () => {
    jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve(''));

    await command.action(logger, { options: { authType: 'browser' } });
    assert.strictEqual(auth.service.authType, AuthType.Browser, 'Incorrect authType set');
  });

  it('correctly handles error when clearing persisted auth information',
    async () => {
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve('ABC'));
      jestUtil.restore(auth.clearConnectionInfo);
      jest.spyOn(auth, 'clearConnectionInfo').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));

      try {
        await command.action(logger, { options: {} });
      }
      finally {
        jestUtil.restore([
          auth.clearConnectionInfo
        ]);
      }
    }
  );

  it('correctly handles error when clearing persisted auth information (debug)',
    async () => {
      jest.spyOn(auth, 'ensureAccessToken').mockClear().mockImplementation(() => Promise.resolve('ABC'));
      jestUtil.restore(auth.clearConnectionInfo);
      jest.spyOn(auth, 'clearConnectionInfo').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));

      try {
        await command.action(logger, { options: { debug: true } });
      }
      finally {
        jestUtil.restore([
          auth.clearConnectionInfo
        ]);
      }
    }
  );

  it('correctly handles error when restoring auth information', async () => {
    jestUtil.restore(auth.restoreAuth);
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));
    try {
      await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('An error has occurred'));
    }
    finally {
      jestUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('fails validation if authType is set to secret and secret option is not specified',
    async () => {
      const actual = await command.validate({ options: { authType: 'secret' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );
});
