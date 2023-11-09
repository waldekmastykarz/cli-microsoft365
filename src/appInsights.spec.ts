import assert from 'assert';
import fs from 'fs';
import { pid } from './utils/pid.js';
import { session } from './utils/session.js';
import { jestUtil } from './utils/jestUtil.js';

const env = Object.assign({}, process.env);

describe('appInsights', () => {
  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      pid.getProcessName,
      session.getId
    ]);
    process.env = env;
  });

  it('adds -dev label to version logged in the telemetry when CLI ran locally',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      const i: any = await import(`./appInsights.js#${Math.random()}`);
      assert(i.default.commonProperties.version.indexOf('-dev') > -1);
    }
  );

  it('doesn\'t add -dev label to version logged in the telemetry when CLI installed from npm',
    async () => {
      const originalExistsSync = fs.existsSync;
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(path => {
        if (path.toString().endsWith('src')) {
          return false;
        }
        else {
          return originalExistsSync(path);
        }
      });
      const i: any = await import(`./appInsights.js#${Math.random()}`);
      assert(i.default.commonProperties.version.indexOf('-dev') === -1);
    }
  );

  it('sets env logged in the telemetry to \'docker\' when CLI run in CLI docker image',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      process.env.CLIMICROSOFT365_ENV = 'docker';
      const i: any = await import(`./appInsights.js#${Math.random()}`);
      assert(i.default.commonProperties.env === 'docker');
    }
  );
});