import assert from 'assert';
import { jestUtil } from './jestUtil.js';

describe('utils/sinonUtil', () => {
  it('doesn\'t fail when restoring stub if the passed object is undefined',
    () => {
      jestUtil.restore(undefined);
      assert(true);
    }
  );
});