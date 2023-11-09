import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../utils/jestUtil.js';
import { tsUtil } from '../../../../../utils/tsUtil.js';
import { TsFile } from './index.js';

describe('TsFile', () => {
  let tsFile: TsFile;

  beforeAll(() => {
    tsFile = new TsFile('foo');
  });

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      tsUtil.createSourceFile
    ]);
    (tsFile as any)._source = undefined;
  });

  it('doesn\'t throw exception if the specified file doesn\'t exist', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    tsFile.source;
    assert(true);
  });

  it('returns undefined source if the specified file doesn\'t exist', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    assert.strictEqual(tsFile.source, undefined);
  });

  it('returns undefined sourceFile if the specified file doesn\'t exist',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      assert.strictEqual(tsFile.sourceFile, undefined);
    }
  );

  it('returns undefined nodes if the specified file doesn\'t exist', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    assert.strictEqual(tsFile.nodes, undefined);
  });

  it('doesn\'t fail when creating TS file fails', () => {
    (tsFile as any)._source = '123';
    jest.spyOn(tsUtil, 'createSourceFile').mockClear().mockImplementation(() => { throw new Error('An exception has occurred'); });
    assert.strictEqual(tsFile.sourceFile, undefined);
  });
});
