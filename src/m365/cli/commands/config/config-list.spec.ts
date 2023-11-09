import assert from 'assert';
import { telemetry } from '../../../../telemetry.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './config-list.js';

describe(commands.CONFIG_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
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
    loggerSpy = jest.spyOn(logger, 'log').mockClear();
  });

  afterEach(() => {
    jestUtil.restore(Cli.getInstance().config.all);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONFIG_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('returns a list of all the self set properties', async () => {
    const config = Cli.getInstance().config;
    jest.spyOn(config, 'all').mockClear().mockImplementation().value({ 'errorOutput': 'stdout' });

    await command.action(logger, { options: {} });
    assert(loggerSpy.calledWith({ 'errorOutput': 'stdout' }));
  });
});
