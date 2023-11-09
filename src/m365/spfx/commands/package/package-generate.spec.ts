import assert from 'assert';
import fs from 'fs';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './package-generate.js';

const admZipMock = {
  // we need these unused params so that they can be properly mocked with sinon
  /* eslint-disable @typescript-eslint/no-unused-vars */
  addFile: (entryName: string, data: Buffer, comment?: string, attr?: number) => { },
  addLocalFile: (localPath: string, zipPath?: string, zipName?: string) => { },
  writeZip: (targetFileName?: string, callback?: (error: Error | null) => void) => { }
  /* eslint-enable @typescript-eslint/no-unused-vars */
};

describe(commands.PACKAGE_GENERATE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    (command as any).archive = admZipMock;
    commandInfo = Cli.getCommandInfo(command);
    Cli.getInstance().config;
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
    jest.spyOn(fs, 'mkdtempSync').mockClear().mockImplementation(_ => '/tmp/abc');
    jest.spyOn(fsUtil, 'readdirR').mockClear().mockImplementation(_ => ['file1.png', 'file.json']);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => 'abc');
    jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    jest.spyOn(fs, 'rmdirSync').mockClear().mockImplementation(_ => { });
    jest.spyOn(fs, 'mkdirSync').mockClear().mockImplementation(_ => '/tmp/abc/def');
    jest.spyOn(fs, 'copyFileSync').mockClear().mockImplementation(_ => { });
    jest.spyOn(fs, 'statSync').mockClear().mockImplementation(src => {
      return {
        isDirectory: () => src.toString().indexOf('.') < 0
      } as any;
    });
  });

  afterEach(() => {
    jestUtil.restore([
      (command as any).generateNewId,
      admZipMock.addFile,
      admZipMock.addLocalFile,
      admZipMock.writeZip,
      fs.copyFileSync,
      fs.mkdtempSync,
      fs.mkdirSync,
      fs.readFileSync,
      fs.rmdirSync,
      fs.statSync,
      fs.writeFileSync,
      fsUtil.copyRecursiveSync,
      fsUtil.readdirR
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PACKAGE_GENERATE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a package for the specified HTML snippet', async () => {
    const archiveWriteZipSpy = jest.spyOn(admZipMock, 'writeZip').mockClear();
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    });
    assert(archiveWriteZipSpy.called);
  });

  it('creates a package for the specified HTML snippet (debug)', async () => {
    const archiveWriteZipSpy = jest.spyOn(admZipMock, 'writeZip').mockClear();
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: true
      }
    });
    assert(archiveWriteZipSpy.called);
  });

  it('creates a package exposed as a Teams tab', async () => {
    jestUtil.restore([fs.readFileSync, fs.writeFileSync]);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'tab'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsTab']).replace(/"/g, '&quot;')));
  });

  it('creates a package exposed as a Teams personal app', async () => {
    jestUtil.restore([fs.readFileSync, fs.writeFileSync]);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'personalApp'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsPersonalApp']).replace(/"/g, '&quot;')));
  });

  it('creates a package exposed as a Teams tab and personal app', async () => {
    jestUtil.restore([fs.readFileSync, fs.writeFileSync]);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsTab', 'TeamsPersonalApp']).replace(/"/g, '&quot;')));
  });

  it('handles exception when creating a temp folder failed', async () => {
    jestUtil.restore(fs.mkdtempSync);
    jest.spyOn(fs, 'mkdtempSync').mockClear().mockImplementation().throws(new Error('An error has occurred'));
    const archiveWriteZipSpy = jest.spyOn(admZipMock, 'writeZip').mockClear();
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }), (err) => err === 'An error has occurred');
    assert(archiveWriteZipSpy.notCalled);
  });

  it('handles error when creating the package failed', async () => {
    jest.spyOn(admZipMock, 'writeZip').mockClear().mockImplementation().throws(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }), (err) => err === 'An error has occurred');
  });

  it('removes the temp directory after the package has been created',
    async () => {
      jestUtil.restore(fs.rmdirSync);
      const fsrmdirSyncSpy = jest.spyOn(fs, 'rmdirSync').mockClear().mockImplementation(_ => { });
      await command.action(logger, {
        options: {
          webPartTitle: 'Amsterdam weather',
          webPartDescription: 'Shows weather in Amsterdam',
          name: 'amsterdam-weather',
          html: 'abc',
          allowTenantWideDeployment: true,
          enableForTeams: 'all'
        }
      });
      assert(fsrmdirSyncSpy.called);
    }
  );

  it('removes the temp directory if creating the package failed', async () => {
    jestUtil.restore(fs.rmdirSync);
    const fsrmdirSyncSpy = jest.spyOn(fs, 'rmdirSync').mockClear().mockImplementation(_ => { });
    jest.spyOn(admZipMock, 'writeZip').mockClear().mockImplementation().throws(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }));
    assert(fsrmdirSyncSpy.called);
  });

  it('prompts user to remove the temp directory manually if removing it automatically failed',
    async () => {
      jestUtil.restore(fs.rmdirSync);
      jest.spyOn(fs, 'rmdirSync').mockClear().mockImplementation().throws(new Error('An error has occurred'));
      await assert.rejects(command.action(logger, {
        options: {
          webPartTitle: 'Amsterdam weather',
          webPartDescription: 'Shows weather in Amsterdam',
          name: 'amsterdam-weather',
          html: 'abc',
          allowTenantWideDeployment: true,
          enableForTeams: 'all'
        }
      }), (err) => err === 'An error has occurred while removing the temp folder at /tmp/abc. Please remove it manually.');
    }
  );

  it('leaves unknown token as-is', async () => {
    jestUtil.restore([fs.readFileSync, fs.writeFileSync]);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => '$token$');
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'tab'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', '$token$'));
  });

  it('exposes page context globally', async () => {
    jestUtil.restore([fs.readFileSync, fs.writeFileSync]);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => '$exposePageContextGlobally$');
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        exposePageContextGlobally: true
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', '!0'));
  });

  it('exposes Teams context globally', async () => {
    jestUtil.restore([fs.readFileSync, fs.writeFileSync]);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => '$exposeTeamsContextGlobally$');
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        exposeTeamsContextGlobally: true
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', '!0'));
  });

  it(`fails validation if the enableForTeams option is invalid`, async () => {
    const actual = await command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to tab`,
    async () => {
      const actual = await command.validate({
        options: {
          webPartTitle: 'Amsterdam weather',
          webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
          html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
          enableForTeams: 'tab'
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it(`passes validation if the enableForTeams option is set to personalApp`,
    async () => {
      const actual = await command.validate({
        options: {
          webPartTitle: 'Amsterdam weather',
          webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
          html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
          enableForTeams: 'personalApp'
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it(`passes validation if the enableForTeams option is set to all`,
    async () => {
      const actual = await command.validate({
        options: {
          webPartTitle: 'Amsterdam weather',
          webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
          html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
          enableForTeams: 'all'
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
