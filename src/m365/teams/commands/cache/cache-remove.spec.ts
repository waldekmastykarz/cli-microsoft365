import assert from 'assert';
import fs from 'fs';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './cache-remove.js';

describe(commands.CACHE_REMOVE, () => {
  const processOutput = `ProcessId
  6456
  14196
  11352`;
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    jest.spyOn(Cli.getInstance().config, 'all').mockClear().mockImplementation().value({});
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

    promptOptions = undefined;

    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: true };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      Cli.prompt,
      (command as any).exec,
      (process as any).kill
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CACHE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before clear cache when confirm option not passed', async () => {
    jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
    jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });

    await command.action(logger, {
      options: {}
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('fails validation if called from docker container.', async () => {
    jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
    jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': 'docker' });

    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if not called from win32 or darwin platform.',
    async () => {
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('android');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      const actual = await command.validate({
        options: {}
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if called from win32 or darwin platform.',
    async () => {
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      const actual = await command.validate({
        options: {}
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails to remove teams cache when exec fails randomly when killing teams.exe process',
    async () => {
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      const error = new Error('random error');
      jest.spyOn(command, 'exec' as any).mockClear().mockImplementation(async (opts) => {
        if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
          throw error;
        }
        throw 'Invalid request';
      });
      await assert.rejects(command.action(logger, { options: { force: true } } as any), new CommandError('random error'));
    }
  );

  it('fails to remove teams cache when exec fails randomly when removing cache folder',
    async () => {
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming' });
      jest.spyOn(process, 'kill' as any).mockClear().mockReturnValue(null);
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      const error = new Error('random error');
      jest.spyOn(command, 'exec' as any).mockClear().mockImplementation(async (opts) => {
        if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
          return { stdout: processOutput };
        }
        if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Roaming\\Microsoft\\Teams"') {
          throw error;
        }
        throw 'Invalid request';
      });
      await assert.rejects(command.action(logger, { options: { force: true } } as any), new CommandError('random error'));
    }
  );

  it('removes Teams cache from macOs platform without prompting.',
    async () => {
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('darwin');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });
      jest.spyOn(command, 'exec' as any).mockClear().mockReturnValue({ stdout: '' });
      jest.spyOn(process, 'kill' as any).mockClear().mockReturnValue(null);
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);

      await command.action(logger, {
        options: {
          force: true,
          verbose: true
        }
      });
      assert(true);
    }
  );

  it('removes teams cache when teams is currently not active', async () => {
    jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
    jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming' });
    jest.spyOn(process, 'kill' as any).mockClear().mockReturnValue(null);
    jest.spyOn(command, 'exec' as any).mockClear().mockImplementation(async (opts) => {
      if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
        return { stdout: 'No Instance(s) Available.' };
      }
      if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Roaming\\Microsoft\\Teams"') {
        return;
      }
      throw 'Invalid request';
    });
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);

    await command.action(logger, {
      options: {
        force: true,
        verbose: true
      }
    });
    assert(true);
  });

  it('removes Teams cache from win32 platform without prompting.',
    async () => {
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('win32');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming' });
      jest.spyOn(process, 'kill' as any).mockClear().mockReturnValue(null);
      jest.spyOn(command, 'exec' as any).mockClear().mockImplementation(async (opts) => {
        if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
          return { stdout: processOutput };
        }
        if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Roaming\\Microsoft\\Teams"') {
          return;
        }
        throw 'Invalid request';
      });
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      await command.action(logger, {
        options: {
          force: true,
          verbose: true
        }
      });
      assert(true);
    }
  );

  it('removes Teams cache from darwin platform with prompting.', async () => {
    jest.spyOn(process, 'platform').mockClear().mockImplementation().value('darwin');
    jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });
    jest.spyOn(command, 'exec' as any).mockClear().mockReturnValue({ stdout: 'pid' });
    jest.spyOn(process, 'kill' as any).mockClear().mockReturnValue(null);
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(true);
  });

  it('aborts cache clearing when no cache folder is found', async () => {
    jest.spyOn(process, 'platform').mockClear().mockImplementation().value('darwin');
    jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
    await command.action(logger, {
      options: {
        verbose: true
      }
    });
  });

  it('aborts cache clearing from Teams when prompt not confirmed',
    async () => {
      const execStub = jest.spyOn(command, 'exec' as any).mockClear().mockImplementation();
      jest.spyOn(process, 'platform').mockClear().mockImplementation().value('darwin');
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

      await command.action(logger, { options: {} });
      assert(execStub.notCalled);
    }
  );
});