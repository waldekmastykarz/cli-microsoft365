import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN014005_CODE_settingsfile } from './FN014005_CODE_settingsfile.js';

describe('FN014005_CODE_settingsfile', () => {
  let findings: Finding[];
  let rule: FN014005_CODE_settingsfile;
  afterEach(() => {
    jestUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014005_CODE_settingsfile();
  });

  it('doesn\'t return notifications if vscode settings file is present',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      const project: Project = {
        path: '/usr/tmp'
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );

  it('returns notifications if vscode settings file is absent', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
