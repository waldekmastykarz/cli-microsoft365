import { ISerializableTokenCache, TokenCacheContext } from '@azure/msal-node';
import assert from 'assert';
import { jestUtil } from '../utils/jestUtil.js';
import { msalCachePlugin } from './msalCachePlugin.js';

const mockCache: ISerializableTokenCache = {
  deserialize: () => { },
  serialize: () => ''
};
const mockCacheContext = new TokenCacheContext(mockCache, false);

describe('msalCachePlugin', () => {
  let mockCacheDeserializeSpy: jest.SpyInstance;
  let mockCacheSerializeSpy: jest.SpyInstance;

  beforeAll(() => {
    mockCacheDeserializeSpy = jest.spyOn(mockCache, 'deserialize').mockClear();
    mockCacheSerializeSpy = jest.spyOn(mockCache, 'serialize').mockClear();
  });

  afterEach(() => {
    mockCacheDeserializeSpy.mockReset();
    mockCacheSerializeSpy.mockReset();
    mockCacheContext.hasChanged = false;
    jestUtil.restore([
      (msalCachePlugin as any).fileTokenStorage.get,
      (msalCachePlugin as any).fileTokenStorage.set
    ]);
  });

  it(`restores token cache from the cache storage`, (done) => {
    jest.spyOn((msalCachePlugin as any).fileTokenStorage, 'get').mockClear().mockImplementation(() => '');
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.called);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't fail restoring cache if cache file not found`, (done) => {
    jest.spyOn((msalCachePlugin as any).fileTokenStorage, 'get').mockClear().mockImplementation(() => Promise.reject('File not found'));
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't fail restoring cache if an error has occurred`, (done) => {
    jest.spyOn((msalCachePlugin as any).fileTokenStorage, 'get').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));
    msalCachePlugin
      .beforeCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheDeserializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`persists cache on disk when cache changed`, (done) => {
    jest.spyOn((msalCachePlugin as any).fileTokenStorage, 'set').mockClear().mockImplementation(() => Promise.resolve());
    mockCacheContext.hasChanged = true;
    msalCachePlugin
      .afterCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheSerializeSpy.called);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't persist cache on disk when cache not changed`, (done) => {
    jest.spyOn((msalCachePlugin as any).fileTokenStorage, 'set').mockClear().mockImplementation(() => Promise.resolve());
    msalCachePlugin
      .afterCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheSerializeSpy.notCalled);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });

  it(`doesn't throw exception when persisting cache failed`, (done) => {
    jest.spyOn((msalCachePlugin as any).fileTokenStorage, 'set').mockClear().mockImplementation(() => Promise.reject('An error has occurred'));
    mockCacheContext.hasChanged = true;
    msalCachePlugin
      .afterCacheAccess(mockCacheContext)
      .then(() => {
        try {
          assert(mockCacheSerializeSpy.called);
          done();
        }
        catch (ex) {
          done(ex);
        }
      }, ex => done(ex));
  });
});