import assert from 'assert';
import fs from 'fs';
import request from '../../../../../../request.js';
import { jestUtil } from '../../../../../../utils/jestUtil.js';
import { Project } from '../../project-model/index.js';
import { DynamicRule } from './DynamicRule.js';

describe('DynamicRule', () => {
  let rule: DynamicRule;

  beforeEach(() => {
    rule = new DynamicRule();
  });

  afterEach(() => {
    jestUtil.restore([
      fs.readFileSync,
      request.head,
      request.post
    ]);
  });

  it(`doesn't return anything if package.json is missing`, async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: undefined
    };
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });

  it(`doesn't return anything if project has no dependencies`, async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {}
    };
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });

  it('returns something is package.json is here', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects();
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('doesnt return anything is package is unsupported', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/sp-taxonomy': '1.3.5',
          '@pnp/sp-clientsvc': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('@pnp/sp-taxonomy/package.json') || path.toString().endsWith('@pnp/sp-clientsvc/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects();
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('doesn\'t return anything if both module and main are missing',
    async () => {
      const project: Project = {
        path: '/usr/tmp',
        packageJson: {
          dependencies: {
            '@pnp/pnpjs': '1.3.5'
          }
        }
      };
      const originalReadFileSync = fs.readFileSync;
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
        if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
          return JSON.stringify({
          });
        }
        else {
          return originalReadFileSync(path);
        }
      });
      jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
      const findings = await rule.visit(project);
      assert.strictEqual(findings.entries.length, 0);
    }
  );

  it('doesn\'t return anything if file is not present on CDN', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().rejects();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'UMD' }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('doesn\'t return anything if module type is not supported', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'CommonJs' }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('adds missing file extension', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle",
          module: "./dist/pnpjs.es5.umd.bundle.min"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'UMD' }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('uses exports from API', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'UMD', exports: ['pnpjs'] }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
    assert.strictEqual(findings.entries[0].globalName, 'pnpjs');
  });
  it('considers all package entries', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js",
          es2015: "./dist/pnpjs.es5.umd.bundle.min.js",
          jspm: {
            main: "./dist/pnpjs.es5.umd.bundle.min.js",
            files: ["./dist/pnpjs.es5.umd.bundle.min.js"]
          },
          spm: {
            main: "./dist/pnpjs.es5.umd.bundle.min.js"
          }
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'UMD' }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('doesnt return anything if package json is missing', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        throw new Error('file doesnt exist');
      }
      else {
        return originalReadFileSync(path);
      }
    });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 0);
  });
  it('returns something for es2015 modules', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'ES2015' }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
  it('returns something for AMD modules', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalReadFileSync = fs.readFileSync;
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path) => {
      if (path.toString().endsWith('@pnp/pnpjs/package.json')) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else {
        return originalReadFileSync(path);
      }
    });
    jest.spyOn(request, 'head').mockClear().mockImplementation().resolves();
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => { return { scriptType: 'AMD' }; });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.entries.length, 1);
  });
});
