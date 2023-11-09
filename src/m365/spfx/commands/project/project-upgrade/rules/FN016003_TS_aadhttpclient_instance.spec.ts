import assert from 'assert';
import fs from 'fs';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN016003_TS_aadhttpclient_instance } from './FN016003_TS_aadhttpclient_instance.js';
import { TsRule } from './TsRule.js';

describe('FN016003_TS_aadhttpclient_instance', () => {
  let findings: Finding[];
  let rule: FN016003_TS_aadhttpclient_instance;
  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016003_TS_aadhttpclient_instance();
  });

  it('returns empty resolution by default', () => {
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

  it('doesn\'t return notifications if AadHttpClient not assigned to a variable',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `new AadHttpClient();`);
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

  it('uses a comment as resource when AadHttpClient created with one argument',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => `const client = new AadHttpClient(this.context.serviceScope);`);
      const project: Project = {
        path: '/usr/tmp',
        tsFiles: [
          new TsFile('foo')
        ]
      };
      rule.visit(project, findings);
      assert(findings[0].occurrences[0].resolution.indexOf('/* your resource */') > -1);
    }
  );
});
