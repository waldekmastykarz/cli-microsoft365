import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN018004_TEAMS_tab96x96_png } from './FN018004_TEAMS_tab96x96_png.js';

describe('FN018004_TEAMS_tab96x96_png', () => {
  let findings: Finding[];
  let rule: FN018004_TEAMS_tab96x96_png;
  afterEach(() => {
    jestUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN018004_TEAMS_tab96x96_png();
  });

  it('returns empty resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('returns empty file name by default', () => {
    assert.strictEqual(rule.file, '');
  });

  it('doesn\'t return notifications if no manifests are present', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if the icon exists', () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        id: 'c93e90e5-6222-45c6-b241-995df0029e3c',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns path to icon with the specified name when fixed name used',
    () => {
      rule = new FN018004_TEAMS_tab96x96_png('tab96x96.png');
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      const project: Project = {
        path: '/usr/tmp',
        manifests: [{
          $schema: 'schema',
          id: 'c93e90e5-6222-45c6-b241-995df0029e3c',
          componentType: 'WebPart',
          path: '/usr/tmp/webpart'
        }]
      };
      rule.visit(project, findings);
      assert.strictEqual(findings[0].occurrences[0].file, path.join('teams', 'tab96x96.png'));
    }
  );

  it('returns path to icon with name following web part ID when no fixed name specified',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      const project: Project = {
        path: '/usr/tmp',
        manifests: [{
          $schema: 'schema',
          id: 'c93e90e5-6222-45c6-b241-995df0029e3c',
          componentType: 'WebPart',
          path: '/usr/tmp/webpart'
        }]
      };
      rule.visit(project, findings);
      assert.strictEqual(findings[0].occurrences[0].file, path.join('teams', 'c93e90e5-6222-45c6-b241-995df0029e3c_color.png'));
    }
  );

  it(`doesn't return notification when web part ID not specified`, () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [{
        $schema: 'schema',
        componentType: 'WebPart',
        path: '/usr/tmp/webpart'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
