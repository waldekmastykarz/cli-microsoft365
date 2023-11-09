import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../utils/jestUtil.js';
import { ScssFile } from './ScssFile.js';

describe('ScssFile', () => {
  afterEach(() => {
    jestUtil.restore([
      fs.readFileSync
    ]);
  });

  it('doesn\'t fail when reading file contents fails', () => {
    const scssFile = new ScssFile('file.scss');
    assert.strictEqual(scssFile.source, undefined);
  });
});
