import assert from 'assert';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './config-get.js';

describe(commands.CONFIG_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    commandInfo = Cli.getCommandInfo(command);
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
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
    jestUtil.restore(Cli.getInstance().config.get);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`gets value of the specified property`, async () => {
    const config = Cli.getInstance().config;
    jest.spyOn(config, 'get').mockClear().mockImplementation(_ => 'json');
    await command.action(logger, { options: { key: settingsNames.output } });
    assert(loggerSpy.calledWith('json'));
  });

  it(`returns undefined if the specified setting is not configured`,
    async () => {
      const config = Cli.getInstance().config;
      jest.spyOn(config, 'get').mockClear().mockImplementation(_ => undefined);
      await command.action(logger, { options: { key: settingsNames.output } });
      assert(loggerSpy.calledWith(undefined));
    }
  );

  it('supports specifying key', () => {
    const options = command.options;
    let containsOptionKey = false;
    options.forEach(o => {
      if (o.option.indexOf('--key') > -1) {
        containsOptionKey = true;
      }
    });
    assert(containsOptionKey);
  });

  it('fails validation if specified key is invalid ', async () => {
    const actual = await command.validate({ options: { key: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure}`,
    async () => {
      const actual = await command.validate({ options: { key: settingsNames.showHelpOnFailure } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
