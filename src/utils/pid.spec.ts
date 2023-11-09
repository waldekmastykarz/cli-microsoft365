import assert from 'assert';
import child_process from 'child_process';
import fs from 'fs';
import os from 'os';
import { cache } from './cache.js';
import { pid } from './pid.js';
import { jestUtil } from './jestUtil.js';

describe('utils/pid', () => {
  let cacheSetValueStub: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(cache, 'getValue').mockClear().mockReturnValue(undefined);
    cacheSetValueStub = jest.spyOn(cache, 'setValue').mockClear().mockReturnValue(undefined);
  });

  afterEach(() => {
    jestUtil.restore([
      os.platform,
      child_process.spawnSync,
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('retrieves process name on Windows', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
    jest.spyOn(child_process, 'spawnSync').mockClear().mockReturnValue({ stdout: 'pwsh' } as any);

    assert.strictEqual(pid.getProcessName(123), 'pwsh');
  });

  it('retrieves process name on macOS', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('darwin');
    jest.spyOn(child_process, 'spawnSync').mockClear().mockReturnValue({ stdout: '/bin/bash' } as any);

    assert.strictEqual(pid.getProcessName(123), '/bin/bash');
  });

  it('retrieves undefined on macOS when retrieving the process name failed',
    () => {
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('darwin');
      jest.spyOn(child_process, 'spawnSync').mockClear().mockReturnValue({ error: 'An error has occurred' } as any);

      assert.strictEqual(pid.getProcessName(123), undefined);
    }
  );

  it('retrieves process name on Linux', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('linux');
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('(pwsh)');

    assert.strictEqual(pid.getProcessName(123), 'pwsh');
  });

  it(`returns undefined on Linux if the process is not found`, () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('linux');
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined name on other platforms', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('android');

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined when retrieving process name on Windows fails', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
    jest.spyOn(child_process, 'spawnSync').mockClear().mockReturnValue({ error: 'An error has occurred' } as any);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined when extracting process name on Windows', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
    jest.spyOn(child_process, 'spawnSync').mockClear().mockImplementation(command => {
      if (command === 'wmic') {
        return {
          stdout: 'Caption\
pwsh.exe\
'
        };
      }

      return {
        error: 'An error has occurred'
      } as any;
    });

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('stores retrieved process name in cache', () => {
    jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
    jest.spyOn(child_process, 'spawnSync').mockClear().mockReturnValue({ stdout: 'pwsh' } as any);

    pid.getProcessName(123);

    expect(cacheSetValueStub).toHaveBeenCalled();
  });

  it('retrieves process name from cache when available', () => {
    jestUtil.restore(cache.getValue);
    jest.spyOn(cache, 'getValue').mockClear().mockReturnValue('pwsh');
    const osPlatformSpy = jest.spyOn(os, 'platform').mockClear();

    assert.strictEqual(pid.getProcessName(123), 'pwsh');
    expect(osPlatformSpy).not.toHaveBeenCalled();
  });
});