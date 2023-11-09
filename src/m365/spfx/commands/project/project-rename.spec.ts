import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './project-rename.js';

describe(commands.PROJECT_RENAME, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let trackEvent: any;
  let telemetryCommandName: any;
  let writeFileSyncSpy: jest.Mock;
  const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react';

  beforeAll(() => {
    trackEvent = jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation((commandName) => {
      telemetryCommandName = commandName;
    });
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    telemetryCommandName = null;
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    writeFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockReturnValue();
  });

  afterEach(() => {
    jestUtil.restore([
      (command as any).generateNewId,
      (command as any).getProjectRoot,
      (command as any).getProject,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_RENAME);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    await command.action(logger, { options: { newName: 'spfx-react' } });
    assert(trackEvent.called);
  });

  it('logs correct telemetry event', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    await command.action(logger, { options: { newName: 'spfx-react' } });
    assert.strictEqual(telemetryCommandName, commands.PROJECT_RENAME);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(null);

    await assert.rejects(command.action(logger, { options: { newName: 'spfx-react' } } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('updates only the files found and skips other files', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));
    jest.spyOn(command as any, 'getProject').mockClear().mockReturnValue({
      path: projectPath,
      packageJson: {
        dependencies: {},
        name: 'spfx'
      }
    });
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
    await command.action(logger, { options: { newName: 'spfx-react' } } as any);
    assert(writeFileSyncSpy.notCalled);
  });

  it('handles error while updating the files', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));
    jest.spyOn(command as any, 'getProject').mockClear().mockReturnValue({
      path: projectPath,
      packageJson: {
        dependencies: {},
        name: 'spfx'
      }
    });
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => { throw 'error'; });
    await assert.rejects(command.action(logger, { options: { newName: 'spfx-react' } } as any),
      new CommandError('error'));
  });

  it('replaces project name in package.json', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    const replacedContent = `{
  "name": "spfx-react",
  "version": "0.0.1",
  "private": true,
  "engines": {
    "node": ">=0.10.0"
  },
  "scripts": {
    "build": "gulp bundle",
    "clean": "gulp clean",
    "test": "gulp test"
  },
  "dependencies": {
    "react": "16.7.0",
    "react-dom": "16.7.0",
    "@types/react": "16.7.22",
    "@types/react-dom": "16.8.0",
    "office-ui-fabric-react": "6.143.0",
    "@microsoft/sp-core-library": "1.8.2",
    "@microsoft/sp-property-pane": "1.8.2",
    "@microsoft/sp-webpart-base": "1.8.2",
    "@microsoft/sp-lodash-subset": "1.8.2",
    "@microsoft/sp-office-ui-fabric-core": "1.8.2",
    "@types/webpack-env": "1.13.1",
    "@types/es6-promise": "0.0.33"
  },
  "resolutions": {
    "@types/react": "16.7.22"
  },
  "devDependencies": {
    "@microsoft/sp-build-web": "1.8.2",
    "@microsoft/sp-tslint-rules": "1.8.2",
    "@microsoft/sp-module-interfaces": "1.8.2",
    "@microsoft/sp-webpart-workbench": "1.8.2",
    "@microsoft/rush-stack-compiler-2.9": "0.7.7",
    "gulp": "~3.9.1",
    "@types/chai": "3.4.34",
    "@types/mocha": "2.2.38",
    "ajv": "~5.2.2"
  }
}`;

    await command.action(logger, { options: { newName: 'spfx-react', generateNewId: true } } as any);
    assert(writeFileSyncSpy.calledWith(expect.any(String), replacedContent, 'utf-8'));
  });

  it('replaces only project name in .yo-rc.json when --generateNewId is not passed',
    async () => {
      jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

      const replacedContent = `{
    "@microsoft/generator-sharepoint": {
      "version": "1.8.2",
      "libraryName": "spfx-react",
      "libraryId": "da1c365f-1532-4e10-aca2-7a0d29c3245b",
      "environment": "spo",
      "packageManager": "npm",
      "solutionName": "spfx-react",
      "skipFeatureDeployment": false,
      "componentType": "webpart",
      "framework": "react",
      "componentName": "HelloWorld",
      "componentDescription": "HelloWorld",
      "isCreatingSolution": true,
      "isDomainIsolated": false
    }
  }`;

      await command.action(logger, { options: { newName: 'spfx-react' } } as any);
      assert(writeFileSyncSpy.calledWith(expect.any(String), replacedContent, 'utf-8'));
    }
  );

  it('replaces project name and id in .yo-rc.json when --generateNewId is passed',
    async () => {
      jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

      jest.spyOn((command as any), 'generateNewId').mockClear().mockReturnValue('69cb6882-acc1-4148-b059-31ae149ba077');

      const replacedContent = `{
    "@microsoft/generator-sharepoint": {
      "version": "1.8.2",
      "libraryName": "spfx-react",
      "libraryId": "69cb6882-acc1-4148-b059-31ae149ba077",
      "environment": "spo",
      "packageManager": "npm",
      "solutionName": "spfx-react",
      "skipFeatureDeployment": false,
      "componentType": "webpart",
      "framework": "react",
      "componentName": "HelloWorld",
      "componentDescription": "HelloWorld",
      "isCreatingSolution": true,
      "isDomainIsolated": false
    }
  }`;

      await command.action(logger, { options: { newName: 'spfx-react', generateNewId: true, debug: true } } as any);
      assert(writeFileSyncSpy.calledWith(expect.any(String), replacedContent, 'utf-8'));
    }
  );

  it('replaces only project name in package-solution.json when --generateNewId is not passed',
    async () => {
      jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

      const replacedContent = `{
    "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
    "solution": {
      "name": "spfx-react-client-side-solution",
      "id": "da1c365f-1532-4e10-aca2-7a0d29c3245b",
      "version": "1.0.0.0",
      "includeClientSideAssets": true,
      "isDomainIsolated": false
    },
    "paths": {
      "zippedPackage": "solution/spfx-react.sppkg"
    }
  }`;

      await command.action(logger, { options: { newName: 'spfx-react' } } as any);
      assert(writeFileSyncSpy.calledWith(expect.any(String), replacedContent, 'utf-8'));
    }
  );

  it('replaces project name and id in package-solution.json when --generateNewId is passed',
    async () => {
      jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

      jest.spyOn((command as any), 'generateNewId').mockClear().mockReturnValue('69cb6882-acc1-4148-b059-31ae149ba077');

      const replacedContent = `{
    "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
    "solution": {
      "name": "spfx-react-client-side-solution",
      "id": "69cb6882-acc1-4148-b059-31ae149ba077",
      "version": "1.0.0.0",
      "includeClientSideAssets": true,
      "isDomainIsolated": false
    },
    "paths": {
      "zippedPackage": "solution/spfx-react.sppkg"
    }
  }`;

      await command.action(logger, { options: { newName: 'spfx-react', generateNewId: true } } as any);
      assert(writeFileSyncSpy.calledWith(expect.any(String), replacedContent, 'utf-8'));
    }
  );

  it('replaces project name in deploy-azure-storage.json', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    const replacedContent = `{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/deploy-azure-storage.schema.json",
  "workingDir": "./temp/deploy/",
  "account": "<!-- STORAGE ACCOUNT NAME -->",
  "container": "spfx-react",
  "accessKey": "<!-- ACCESS KEY -->"
}`;

    await command.action(logger, { options: { newName: 'spfx-react' } } as any);
    assert(writeFileSyncSpy.calledWith(expect.any(String), replacedContent, 'utf-8'));
  });

  it('replaces project name in README.md', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    let replacedContent = `## spfx-react

This is where you include your WebPart documentation.

### Building the code

\`\`\`bash
git clone the repo
npm i
npm i -g gulp
gulp
\`\`\`

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
`;

    await command.action(logger, { options: { newName: 'spfx-react', debug: true } } as any);
    let fileSyncContent: string = writeFileSyncSpy.mock.lastCall[1];
    fileSyncContent = fileSyncContent.replace(/(\r\n|\n|\r)/gm, "");
    replacedContent = replacedContent.replace(/(\r\n|\n|\r)/gm, "");
    assert.strictEqual(fileSyncContent, replacedContent);
    assert.strictEqual(loggerLogToStderrSpy.mock.calls[5][0], `Updated README.md`);
  });
});
