import assert from 'assert';
import { session } from '../utils/session.js';
import { cache } from './cache.js';
import { jestUtil } from './jestUtil.js';

describe('utils/session', () => {
  afterEach(() => {
    jestUtil.restore([
      cache.getValue,
      cache.setValue
    ]);
  });

  it('returns existing session ID if available', () => {
    jest.spyOn(cache, 'getValue').mockClear().mockImplementation(() => '123');
    assert.strictEqual(session.getId(1), '123');
  });

  it('returns new session ID if no ID available', () => {
    jest.spyOn(cache, 'getValue').mockClear().mockReturnValue(undefined);
    jest.spyOn(cache, 'setValue').mockClear().mockImplementation(() => { });
    assert(session.getId(1).length > 3);
  });
});