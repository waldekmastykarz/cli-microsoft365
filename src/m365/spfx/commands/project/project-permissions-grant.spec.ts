import assert from 'assert';
import chalk from 'chalk';
import fs from 'fs';
import path from 'path';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import spoServicePrincipalGrantAddCommand from '../../../spo/commands/serviceprincipal/serviceprincipal-grant-add.js';
import commands from '../../commands.js';
import command from './project-permissions-grant.js';

describe(commands.PROJECT_PERMISSIONS_GRANT, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let loggerStderrLogSpy: jest.SpyInstance;
  const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react';
  const packagejsonContent = `{
    "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
    "solution": {
      "name": "hello-world-client-side-solution",
      "id": "ae661453-5f05-44f3-8312-66eb8f36f6fc",
      "version": "1.0.0.0",
      "includeClientSideAssets": true,
      "skipFeatureDeployment": true,
      "isDomainIsolated": false,
      "developer": {
        "name": "",
        "websiteUrl": "",
        "privacyUrl": "",
        "termsOfUseUrl": "",
        "mpnId": "Undefined-1.16.1"
      },
      "metadata": {
        "shortDescription": {
          "default": "hello world description"
        },
        "longDescription": {
          "default": "hello world description"
        },
        "screenshotPaths": [],
        "videoUrl": "",
        "categories": []
      },
      "features": [
        {
          "title": "Hello world Feature",
          "description": "The feature that activates elements of the hello world solution.",
          "id": "345cee0f-e4fb-464f-b649-20fc96b5f6aa",
          "version": "1.0.0.0"
        }
      ],
      "webApiPermissionRequests": [
        {
          "resource": "Microsoft Graph",
          "scope": "User.ReadBasic.All"
        }
      ]
    },
    "paths": {
      "zippedPackage": "solution/hello-worldsppkg"
    }
  }`;
  const grantResponse = {
    "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d",
    "ConsentType": "AllPrincipals",
    "IsDomainIsolated": false,
    "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM",
    "PackageName": null,
    "Resource": "Microsoft Graph",
    "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903",
    "Scope": "User.ReadBasic.All"
  };

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    loggerStderrLogSpy = jest.spyOn(logger, 'logToStderr').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      (command as any).getProjectRoot,
      fs.existsSync,
      fs.readFileSync,
      Cli.executeCommandWithOutput
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_PERMISSIONS_GRANT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('handles correctly when the package-solution.json file is not found',
    async () => {
      jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);

      await assert.rejects(command.action(logger, { options: {} } as any),
        new CommandError(`The package-solution.json file could not be found`));
    }
  );

  it('grant the specified permissions from the package-solution.json file',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(packagejsonContent);

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoServicePrincipalGrantAddCommand) {
          return ({
            stdout: `{ "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903", "Scope": "User.ReadBasic.All"}`
          });
        }

        throw new CommandError('Unknown case');
      });

      await command.action(logger, {
        options: {
          debug: true
        }
      });
      assert(loggerLogSpy.calledWith(grantResponse));
    }
  );

  it('shows warning when permission already exist', async () => {
    const grantExistError = {
      error: {
        message: 'An OAuth permission with the resource Microsoft Graph and scope User.ReadBasic.All already exists.Parameter name: permissionRequest'
      },
      stderr: ''
    };

    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(packagejsonContent);

    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === spoServicePrincipalGrantAddCommand) {
        throw grantExistError;
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
      }
    });
    assert.strictEqual(loggerStderrLogSpy.calledWith(chalk.yellow("An OAuth permission with the resource Microsoft Graph and scope User.ReadBasic.All already exists.Parameter name: permissionRequest")), true);
  });

  it('correctly handles error when something went wrong when granting permission',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(packagejsonContent);

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoServicePrincipalGrantAddCommand) {
          throw 'Something went wrong';
        }

        throw new CommandError('Unknown case');
      });

      await assert.rejects(command.action(logger, { options: {} } as any),
        new CommandError(`Something went wrong`));
    }
  );
});
