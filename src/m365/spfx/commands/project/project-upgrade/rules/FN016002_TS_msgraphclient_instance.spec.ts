import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN016002_TS_msgraphclient_instance } from './FN016002_TS_msgraphclient_instance.js';
import { TsRule } from './TsRule.js';

describe('FN016002_TS_msgraphclient_instance', () => {
  let findings: Finding[];
  let rule: FN016002_TS_msgraphclient_instance;
  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016002_TS_msgraphclient_instance();
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

  it('doesn\'t return notifications if couldn\'t retrieve the call expression',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `this.serviceScope.consume;`);
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

  it('doesn\'t return notifications if service key is a constant', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `this.serviceScope.consume("abc");`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if service key is not MSGraphClient',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `this.serviceScope.consume(ContosoClient.serviceKey);`);
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

  it('doesn\'t return notifications if retrieved MSGraphClient is not assigned to a variable',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `this.serviceScope.consume(MSGraphClient.serviceKey);`);
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
