import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN025001_ESLINTRCJS_overrides } from './FN025001_ESLINTRCJS_overrides.js';

describe('FN025001_ESLINTRCJS_overrides', () => {
  let findings: Finding[];
  let rule: FN025001_ESLINTRCJS_overrides;

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  beforeEach(() => {
    rule = new FN025001_ESLINTRCJS_overrides('{ foo: bar }');
    findings = [];
  });

  it('doesn\'t return notification if .eslintrc.js not found', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if .eslintrc.js is found but no nodes are present',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => ``);
      const project: Project = {
        path: '/usr/tmp',
        esLintRcJs: new TsFile('foo')
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );

  it('doesn\'t return notification if .eslintrc.js is found but module is not present',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `foo`);
      const project: Project = {
        path: '/usr/tmp',
        esLintRcJs: new TsFile('foo')
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );

  it('file returned is ./.eslintrc.js when found', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(rule.file, './.eslintrc.js');
  });

  it('doesn\'t return notification if overrides property is present', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `export default { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('does return notification if overrides property is not present', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('returns resolution for finding if overrides property is not present',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname } }`);
      const project: Project = {
        path: '/usr/tmp',
        esLintRcJs: new TsFile('foo')
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 1);
      assert.strictEqual(rule.resolution, 'export default {\n      overrides: [\n        { foo: bar }\n      ]\n    };');
    }
  );

  it('does not return resolution for finding if overrides property is present',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `export default { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { } ] } }`);
      const project: Project = {
        path: '/usr/tmp',
        esLintRcJs: new TsFile('foo')
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );
});
