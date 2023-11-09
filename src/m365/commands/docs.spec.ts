import assert from 'assert';
import { Cli } from '../../cli/Cli.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { app } from '../../utils/app.js';
import { browserUtil } from '../../utils/browserUtil.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import commands from './commands.js';
import command from './docs.js';

describe(commands.DOCS, () => {
  let log: any[];
  let logger: Logger;
  let cli: Cli;
  let loggerLogSpy: jest.SpyInstance;
  let getSettingWithDefaultValueStub: jest.Mock;

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
  });

  beforeEach(() => {
    log = [];
    cli = Cli.getInstance();
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
    getSettingWithDefaultValueStub = jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockReturnValue(false);
  });

  afterEach(() => {
    jestUtil.restore([
      loggerLogSpy,
      getSettingWithDefaultValueStub
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DOCS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should log a message and return if autoOpenLinksInBrowser is false',
    async () => {
      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith(app.packageJson().homepage));
    }
  );

  it('should open the CLI for Microsoft 365 docs webpage URL using "open" if autoOpenLinksInBrowser is true',
    async () => {
      getSettingWithDefaultValueStub.mockRestore();
      getSettingWithDefaultValueStub = jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockReturnValue(true);

      const openStub = jest.spyOn(browserUtil, 'open').mockClear().mockImplementation(async (url) => {
        if (url === 'https://pnp.github.io/cli-microsoft365/') {
          return;
        }
        throw 'Invalid url';
      });
      await command.action(logger, { options: {} });
      assert(openStub.calledWith(app.packageJson().homepage));
    }
  );
});