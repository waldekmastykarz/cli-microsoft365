import assert from 'assert';
import fs from 'fs';
import { spfx } from '../../../../../../utils/spfx.js';
import { Project, ScssFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN022002_SCSS_add_fabric_react } from './FN022002_SCSS_add_fabric_react.js';

describe('FN022002_SCSS_add_fabric_react', () => {
  let findings: Finding[];
  let rule: FN022002_SCSS_add_fabric_react;
  let fileStub: jest.Mock;
  let utilsStub: jest.Mock;

  beforeEach(() => {
    findings = [];
    utilsStub = jest.spyOn(spfx, 'isReactProject').mockClear().mockReturnValue(true);
  });

  afterEach(() => {
    fileStub.mockRestore();
    utilsStub.mockRestore();
  });

  it('doesn\'t return notifications if import is already there', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');

    fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('~fabric-ui/react');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if import is missing and no condition', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notifications if import is missing but condition is not met',
    () => {
      rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react', '~old-fabric-ui/react');

      fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('');
      const project: Project = {
        path: '/usr/tmp',
        scssFiles: [
          new ScssFile('some/path')
        ]
      };
      rule.visit(project, findings);
      assert.strictEqual(findings.length, 0);
    }
  );

  it('returns notifications if import is missing and condition is met', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react', '~old-fabric-ui/react');
    fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('~old-fabric-ui/react');
    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notifications if scss is not in react web part', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    utilsStub.mockRestore();
    utilsStub = jest.spyOn(spfx, 'isReactProject').mockClear().mockReturnValue(false);

    fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: [
        new ScssFile('some/path')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if no scss files', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: []
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('rule file name is empty', () => {
    rule = new FN022002_SCSS_add_fabric_react('~fabric-ui/react');
    fileStub = jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('');

    const project: Project = {
      path: '/usr/tmp',
      scssFiles: []
    };
    rule.visit(project, findings);
    assert.strictEqual(rule.file, '');
  });
});
