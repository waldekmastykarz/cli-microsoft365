import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { cache } from './cache.js';
import { jestUtil } from './jestUtil.js';

describe('utils/cache', () => {
  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.mkdirSync,
      fs.writeFile,
      fs.readdir,
      fs.stat,
      fs.unlink,
      cache.clearExpired
    ]);
  });

  describe('getValue', () => {
    it(`returns undefined when the specified value doesn't exist in cache`,
      () => {
        jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
        assert.strictEqual(cache.getValue('key'), undefined);
      }
    );

    it('returns the specified value from cache', () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('value');
      assert.strictEqual(cache.getValue('key'), 'value');
    });

    it('returns undefined if an error occurs while reading cache', () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation().throws();
      assert.strictEqual(cache.getValue('key'), undefined);
    });

    it('clears expired values', () => {
      const clearExpiredSpy = jest.spyOn(cache, 'clearExpired').mockClear();
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
      cache.getValue('key');

      assert(clearExpiredSpy.called);
    });
  });

  describe('setValue', () => {
    it('clears expired values', () => {
      const clearExpiredSpy = jest.spyOn(cache, 'clearExpired').mockClear();
      jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation().throws();
      cache.setValue('key', 'value');

      assert(clearExpiredSpy.called);
    });

    it(`doesn't fail when creating the cache folder fails`, () => {
      jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation().throws();
      const writeFilesSpy = jest.spyOn(fs, 'writeFile').mockClear();
      cache.setValue('key', 'value');

      assert(writeFilesSpy.notCalled);
    });

    it(`doesn't fail when writing value to cache file fails`, (done) => {
      jest.spyOn(fs, 'mkdirSync').mockClear().mockReturnValue(undefined);
      jest.spyOn(fs, 'writeFile').mockClear().mockReturnValue(done()).mockImplementation((...args: any[]) => args[2]('error'));
      try {
        cache.setValue('key', 'value');
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`writes value to cache in a cache file`, (done) => {
      jest.spyOn(fs, 'mkdirSync').mockClear().mockReturnValue(undefined);
      jest.spyOn(fs, 'writeFile').mockClear().mockReturnValue(done());
      try {
        cache.setValue('key', 'value');
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  describe('clearExpired', () => {
    it(`doesn't fail when reading the cache folder fails`, (done) => {
      jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1]('error'));
      try {
        cache.clearExpired(() => {
          done();
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't fail when the cache folder is empty`, (done) => {
      jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, []));
      try {
        cache.clearExpired(() => {
          done();
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`skips directories while clearing expired entries (dir + file)`,
      (done) => {
        jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['directory', 'file']));
        const twoDaysAgo = new Date();
        twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
        jest.spyOn(fs, 'stat').mockClear().mockImplementation()
          .onFirstCall().mockImplementation((...args: any[]) => args[1](undefined, { isDirectory: () => true }))
          .onSecondCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
          );
        const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
          .mockReturnValue().mockImplementation((...args: any[]) => args[1]());
        try {
          cache.clearExpired(() => {
            try {
              assert(unlinkStub.calledWith(path.join(cache.cacheFolderPath, 'file')));
              done();
            }
            catch (ex) {
              done(ex);
            }
          });
        }
        catch (ex) {
          done(ex);
        }
      }
    );

    it(`skips directories while clearing expired entries (dir only)`, (done) => {
      jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['directory']));
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      jest.spyOn(fs, 'stat').mockClear().mockImplementation()
        .onFirstCall().mockImplementation((...args: any[]) => args[1](undefined, { isDirectory: () => true }));
      const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
        .mockReturnValue().mockImplementation((...args: any[]) => args[1]());
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.notCalled);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't fail while reading file information fails`, (done) => {
      jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['file1', 'file2']));
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      jest.spyOn(fs, 'stat').mockClear().mockImplementation()
        .onFirstCall().mockImplementation((...args: any[]) => args[1]('error'))
        .onSecondCall().mockImplementation(
          (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
        );
      const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
        .mockReturnValue().mockImplementation((...args: any[]) => args[1]());
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledOnce);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't fail while removing expired cache entry fails`, (done) => {
      jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['file']));
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      jest.spyOn(fs, 'stat').mockClear().mockImplementation().mockImplementation(
        (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
      );
      const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
        .mockReturnValue().mockImplementation((...args: any[]) => args[1]('error'));
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledOnce);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't remove cache entries that have been accessed in the last 24 hours`,
      (done) => {
        jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['file1', 'file2']));
        const twoDaysAgo = new Date();
        twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
        jest.spyOn(fs, 'stat').mockClear().mockImplementation()
          .onFirstCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: new Date() })
          )
          .onSecondCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
          );
        const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
          .mockReturnValue().mockImplementation((...args: any[]) => args[1]());
        try {
          cache.clearExpired(() => {
            try {
              assert(unlinkStub.calledOnce);
              done();
            }
            catch (ex) {
              done(ex);
            }
          });
        }
        catch (ex) {
          done(ex);
        }
      }
    );

    it(`doesn't remove cache entries that have been accessed in the last 24 hours (last file recently accessed)`,
      (done) => {
        jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['file1', 'file2']));
        const twoDaysAgo = new Date();
        twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
        jest.spyOn(fs, 'stat').mockClear().mockImplementation()
          .onFirstCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
          )
          .onSecondCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: new Date() })
          );
        const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
          .mockReturnValue().mockImplementation((...args: any[]) => args[1]());
        try {
          cache.clearExpired(() => {
            try {
              assert(unlinkStub.calledOnce);
              done();
            }
            catch (ex) {
              done(ex);
            }
          });
        }
        catch (ex) {
          done(ex);
        }
      }
    );

    it(`removes cache entries that haven't been accessed in the last 24 hours`,
      (done) => {
        jest.spyOn(fs, 'readdir').mockClear().mockImplementation().mockImplementation((...args: any[]) => args[1](undefined, ['file1', 'file2']));
        const twoDaysAgo = new Date();
        twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
        jest.spyOn(fs, 'stat').mockClear().mockImplementation()
          .onFirstCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
          )
          .onSecondCall().mockImplementation(
            (...args: any[]) => args[1](undefined, { isDirectory: () => false, atime: twoDaysAgo })
          );
        const unlinkStub = jest.spyOn(fs, 'unlink').mockClear()
          .mockReturnValue().mockImplementation((...args: any[]) => args[1]());
        try {
          cache.clearExpired(() => {
            try {
              assert(unlinkStub.calledTwice);
              done();
            }
            catch (ex) {
              done(ex);
            }
          });
        }
        catch (ex) {
          done(ex);
        }
      }
    );
  });
});