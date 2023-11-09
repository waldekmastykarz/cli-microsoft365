import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN016001_TS_msgraphclient_packageName } from './FN016001_TS_msgraphclient_packageName.js';
import { TsRule } from './TsRule.js';

describe('FN016001_TS_msgraphclient_packageName', () => {
  let findings: Finding[];
  let rule: FN016001_TS_msgraphclient_packageName;
  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016001_TS_msgraphclient_packageName('@microsoft/sp-http');
  });

  it('returns empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notifications if no .ts files found', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if specified .ts file not found', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if couldn\'t retrieve the import declaration',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `import { MSGraphClient } from '@microsoft/sp-http';`);
      jest.spyOn(TsRule as any, 'getParentOfType').mockClear().mockImplementation(() => undefined);
      const project: Project = {
        path: '/usr/tmp',
        tsFiles: [
          new TsFile('foo')
        ]
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );

  it('doesn\'t return notifications if MSGraphClient is already imported from the correct package',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `import { MSGraphClient } from '@microsoft/sp-http';`);
      const project: Project = {
        path: '/usr/tmp',
        tsFiles: [
          new TsFile('foo')
        ]
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );
});
