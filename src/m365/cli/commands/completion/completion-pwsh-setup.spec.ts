import assert from 'assert';
import fs from 'fs';
import os from 'os';
import path from 'path';
import url from 'url';
import { CommandError } from '../../../../Command.js';
import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './completion-pwsh-setup.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

describe(commands.COMPLETION_PWSH_SETUP, () => {
  const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
    jest.spyOn(autocomplete, 'generateShCompletion').mockClear().mockImplementation(() => { });
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.mkdirSync,
      fs.writeFileSync,
      fs.appendFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_PWSH_SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('appends completion scripts to profile when profile file already exists',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

      await command.action(logger, { options: { profile: profilePath } });
      assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'));
    }
  );

  it('appends completion scripts to profile when profile file already exists (debug)',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

      await command.action(logger, { options: { debug: true, profile: profilePath } });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('creates profile file when it does not exist and appends the completion script to it',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation((path) => path.toString().indexOf('.ps1') < 0);
      const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

      await command.action(logger, { options: { profile: profilePath } });
      assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
      assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
    }
  );

  it('creates profile file when it does not exist and appends the completion script to it (debug)',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation((path) => path.toString().indexOf('.ps1') < 0);
      const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

      await command.action(logger, { options: { debug: true, profile: profilePath } });
      assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
      assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
    }
  );

  it('creates profile path when it does not exist and appends the completion script to it',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      const mkdirSyncStub: jest.Mock = jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation(_ => '');
      const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

      await command.action(logger, { options: { profile: profilePath } });
      assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
      assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
      assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
    }
  );

  it('creates profile path when it does not exist and appends the completion script to it (debug)',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      const mkdirSyncStub: jest.Mock = jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation(_ => '');
      const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

      await command.action(logger, { options: { debug: true, profile: profilePath } });
      assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
      assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
      assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
    }
  );

  it('handles exception when creating profile path', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    const mkdirSyncStub: jest.Mock = jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation(() => { throw error; });
    const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
    const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

    await assert.rejects(command.action(logger, { options: { profile: profilePath } } as any), new CommandError(error));
    assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
    assert(writeFileSyncStub.notCalled, 'Profile file created');
    assert(appendFileSyncStub.notCalled, 'Completion script appended');
  });

  it('handles exception when creating profile file', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { throw error; });
    const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { });

    await assert.rejects(command.action(logger, { options: { profile: profilePath } } as any), new CommandError(error));
    assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
    assert(appendFileSyncStub.notCalled, 'Completion script appended');
  });

  it('handles exception when appending completion script to the profile file',
    async () => {
      const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
      const error: string = 'Unexpected error';
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { });
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      const appendFileSyncStub: jest.Mock = jest.spyOn(fs, 'appendFileSync').mockClear().mockImplementation(() => { throw error; });

      await assert.rejects(command.action(logger, { options: { profile: profilePath } } as any), new CommandError(error));
      assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
    }
  );
});
