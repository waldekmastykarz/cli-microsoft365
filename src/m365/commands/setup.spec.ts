import assert from 'assert';
import { Cli } from '../../cli/Cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { CheckStatus, formatting } from '../../utils/formatting.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import commands from './commands.js';
import command, { SettingNames } from './setup.js';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets.js';

describe(commands.SETUP, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    (command as any).answers = {};
  });

  afterEach(() => {
    jestUtil.restore([
      (command as any).configureSettings,
      Cli.getInstance().config.set,
      pid.isPowerShell
    ]);
  });

  afterAll(() => {
    jestUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets correct settings for interactive, beginner', async () => {
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Beginner',
      summary: true
    };
    const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for interactive, proficient', async () => {
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Proficient',
      summary: true
    };
    const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'options';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, non-PowerShell, beginner',
    async () => {
      (command as any).answers = {
        usageMode: 'Scripting',
        usedInPowerShell: false,
        experience: 'Beginner',
        summary: true
      };
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

      const expected: SettingNames = {};
      Object.assign(expected, scriptingPreset);
      expected.helpMode = 'full';
      (command as any).settings = expected;

      await command.action(logger, { options: {} });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it('sets correct settings for scripting, PowerShell, beginner', async () => {
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: true,
      experience: 'Beginner',
      summary: true
    };
    const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, scriptingPreset);
    Object.assign(expected, powerShellPreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.calledWith(expected));
  });

  it('sets correct settings for scripting, non-PowerShell, proficient',
    async () => {
      (command as any).answers = {
        usageMode: 'Scripting',
        usedInPowerShell: false,
        experience: 'Proficient',
        summary: true
      };
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

      const expected: SettingNames = {};
      Object.assign(expected, scriptingPreset);
      expected.helpMode = 'options';
      (command as any).settings = expected;

      await command.action(logger, { options: {} });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it('sets correct settings for scripting, PowerShell, proficient',
    async () => {
      (command as any).answers = {
        usageMode: 'Scripting',
        usedInPowerShell: true,
        experience: 'Proficient',
        summary: true
      };
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

      const expected: SettingNames = {};
      Object.assign(expected, scriptingPreset);
      Object.assign(expected, powerShellPreset);
      expected.helpMode = 'options';
      (command as any).settings = expected;

      await command.action(logger, { options: {} });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it(`doesn't apply settings when not confirmed`, async () => {
    (command as any).answers = {
      usageMode: 'Scripting',
      usedInPowerShell: false,
      experience: 'Beginner',
      summary: false
    };
    const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

    await command.action(logger, { options: {} });

    assert(configureSettingsStub.notCalled);
  });

  it('sets correct settings for interactive, non-PowerShell via option',
    async () => {
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

      const expected: SettingNames = {};
      Object.assign(expected, interactivePreset);
      (command as any).settings = expected;

      await command.action(logger, { options: { interactive: true } });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it('sets correct settings for scripting, non-PowerShell via option',
    async () => {
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });

      const expected: SettingNames = {};
      Object.assign(expected, scriptingPreset);
      (command as any).settings = expected;

      await command.action(logger, { options: { scripting: true } });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it('sets correct settings for interactive, PowerShell via option',
    async () => {
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });
      jest.spyOn(pid, 'isPowerShell').mockClear().mockImplementation(() => true);

      const expected: SettingNames = {};
      Object.assign(expected, interactivePreset);
      Object.assign(expected, powerShellPreset);
      (command as any).settings = expected;

      await command.action(logger, { options: { interactive: true } });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it('sets correct settings for scripting, PowerShell via option',
    async () => {
      const configureSettingsStub = jest.spyOn(command as any, 'configureSettings').mockClear().mockImplementation(() => { });
      jest.spyOn(pid, 'isPowerShell').mockClear().mockImplementation(() => true);

      const expected: SettingNames = {};
      Object.assign(expected, scriptingPreset);
      Object.assign(expected, powerShellPreset);
      (command as any).settings = expected;

      await command.action(logger, { options: { scripting: true } });

      assert(configureSettingsStub.calledWith(expected));
    }
  );

  it('outputs settings to configure to console in debug mode', async () => {
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Beginner',
      summary: true
    };
    jest.spyOn(Cli.getInstance().config, 'set').mockClear().mockImplementation(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: { debug: true } });

    assert(loggerLogToStderrSpy.calledWith(JSON.stringify(expected, null, 2)));
  });

  it('logs configured settings when used interactively', async () => {
    (command as any).answers = {
      usageMode: 'Interactively',
      experience: 'Beginner',
      summary: true
    };
    jest.spyOn(Cli.getInstance().config, 'set').mockClear().mockImplementation(() => { });

    const expected: SettingNames = {};
    Object.assign(expected, interactivePreset);
    expected.helpMode = 'full';
    (command as any).settings = expected;

    await command.action(logger, { options: {} });

    for (const [key, value] of Object.entries(expected)) {
      assert(loggerLogToStderrSpy.calledWith(formatting.getStatus(CheckStatus.Success, `${key}: ${value}`)), `Expected ${key} to be set to ${value}`);
    }
  });

  it('in the confirmation message lists all settings and their values',
    async () => {
      const settings: SettingNames = {};
      Object.assign(settings, interactivePreset);
      settings.helpMode = 'full';
      const actual = (command as any).getSummaryMessage(settings);

      for (const [key, value] of Object.entries(settings)) {
        assert(actual.indexOf(`- ${key}: ${value}`) > -1, `Expected ${key} to be set to ${value}`);
      }
    }
  );

  it('fails validation when both interactive and scripting options specified',
    async () => {
      const actual = await command.validate({
        options: {
          interactive: true,
          scripting: true
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when no options specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when interactive option specified', async () => {
    const actual = await command.validate({
      options: {
        interactive: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when scripting option specified', async () => {
    const actual = await command.validate({
      options: {
        scripting: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
