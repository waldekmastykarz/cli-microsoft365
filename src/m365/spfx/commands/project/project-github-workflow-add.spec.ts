import assert from 'assert';
import fs from 'fs';
import path from 'path';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './project-github-workflow-add.js';

describe(commands.PROJECT_GITHUB_WORKFLOW_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const projectPath: string = 'test-project';

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    jestUtil.restore([
      (command as any).getProjectRoot,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_GITHUB_WORKFLOW_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if loginMethod is not valid type', async () => {
    const actual = await command.validate({ options: { loginMethod: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is not valid type', async () => {
    const actual = await command.validate({ options: { scope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is sitecollection but the siteUrl was not defined',
    async () => {
      const actual = await command.validate({ options: { scope: 'sitecollection' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection', siteUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required properties are provided', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/project' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('creates a default workflow (debug)', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation().resolves({});

    await command.action(logger, { options: { debug: true } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.github', 'workflows', 'deploy-spfx-solution.yml')), 'workflow file not created');
  });

  it('creates a default workflow with specifying options', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation((fakePath) => {
      if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('.github') && (options as fs.MakeDirectoryOptions).recursive) {
        return `${projectPath}/.github`;
      }

      return '';
    });

    const writeFileSyncStub: jest.Mock = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation().resolves({});

    await command.action(logger, { options: { name: 'test', branchName: 'dev', manuallyTrigger: true, skipFeatureDeployment: true, overwrite: true, loginMethod: 'user', scope: 'sitecollection' } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.github', 'workflows', 'deploy-spfx-solution.yml')), 'workflow file not created');
  });

  it('handles unexpected error', async () => {
    jest.spyOn(command as any, 'getProjectRoot').mockClear().mockReturnValue(path.join(process.cwd(), projectPath));

    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(() => { throw 'error'; });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('error'));
  });
});