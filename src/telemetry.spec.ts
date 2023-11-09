import assert from 'assert';
import child_process from 'child_process';
import { Cli } from "./cli/Cli.js";
import { settingsNames } from './settingsNames.js';
import { telemetry } from './telemetry.js';
import { pid } from './utils/pid.js';
import { jestUtil } from './utils/jestUtil.js';
import { session } from './utils/session.js';

describe('Telemetry', () => {
  let spawnStub: jest.Mock;
  let stdin: string = '';

  beforeAll(() => {
    spawnStub = jest.spyOn(child_process, 'spawn').mockClear().mockImplementation(() => {
      return {
        stdin: {
          write: (s: string) => {
            stdin += s;
          },
          end: () => { }
        },
        unref: () => { }
      } as any;
    });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => 'abc123');
  });

  afterEach(() => {
    jestUtil.restore([
      Cli.getInstance().getSettingWithDefaultValue,
      (telemetry as any).trackTelemetry
    ]);
    spawnStub.mockReset();
    stdin = '';
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it(`doesn't log an event when disableTelemetry is set`, () => {
    jest.spyOn(Cli.getInstance(), 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    telemetry.trackEvent('foo bar', {});
    assert(spawnStub.notCalled);
  });

  it('logs an event when disableTelemetry is not set', () => {
    jest.spyOn(Cli.getInstance(), 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    telemetry.trackEvent('foo bar', {});
    assert(spawnStub.called);
  });

  it(`doesn't log an exception when disableTelemetry is set`, () => {
    jest.spyOn(Cli.getInstance(), 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return true;
      }

      return defaultValue;
    });
    telemetry.trackException('Error!');
    assert(spawnStub.notCalled);
  });

  it('logs an exception when disableTelemetry is not set', () => {
    jest.spyOn(Cli.getInstance(), 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.disableTelemetry) {
        return false;
      }

      return defaultValue;
    });
    telemetry.trackException('Error!');
    assert(spawnStub.called);
  });

  it(`logs an empty string for shell if it couldn't resolve shell process name`,
    () => {
      jest.spyOn(Cli.getInstance(), 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.disableTelemetry) {
          return false;
        }

        return defaultValue;
      });
      jestUtil.restore(pid.getProcessName);
      jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => undefined);

      telemetry.trackEvent('foo bar', {});
      assert.strictEqual(JSON.parse(stdin).shell, '');
    }
  );

  it(`silently handles exception if an error occurs while spawning telemetry runner`,
    (done) => {
      jest.spyOn(Cli.getInstance(), 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.disableTelemetry) {
          return false;
        }

        return defaultValue;
      });
      jestUtil.restore(child_process.spawn);
      jest.spyOn(child_process, 'spawn').mockClear().mockImplementation().throws();
      try {
        telemetry.trackEvent('foo bar', {});
        done();
      }
      catch (e) {
        done(e);
      }
    }
  );
});