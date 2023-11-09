import assert from 'assert';
import fs from 'fs';
import auth from '../../Auth.js';
import { CommandError } from '../../Command.js';
import { Cli } from '../../cli/Cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import request from '../../request.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import PeriodBasedReport from './PeriodBasedReport.js';

class MockCommand extends PeriodBasedReport {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public get usageEndpoint(): string {
    return 'MockEndPoint';
  }

  public commandHelp(): void {
  }
}

describe('PeriodBasedReport', () => {
  const mockCommand = new MockCommand();
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(mockCommand);
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
    (mockCommand as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      fs.writeFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(mockCommand.name, 'mock');
  });

  it('has a description', () => {
    assert.notStrictEqual(mockCommand.description, null);
  });

  it('fails validation on invalid period', async () => {
    const actual = await mockCommand.validate({ options: { period: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid \'D7\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D7'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'D30\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D30'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'D90\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D90'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'180\' period', async () => {
    const actual = await mockCommand.validate({
      options: {
        period: 'D90'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('get unique device type in teams and export it in a period', async () => {
    const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
        return `
        Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
        2019-08-28,0,0,0,0,0,0,7
        `;
      }

      throw 'Invalid request';
    });

    await mockCommand.action(logger, { options: { period: 'D7' } });
    assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
    assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
  });

  it('produce export using period format and Teams unique device type output in txt',
    async () => {
      const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
          return `
          Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
          2019-08-28,0,0,0,0,0,0,7
          `;
        }

        throw 'Invalid request';
      });

      await mockCommand.action(logger, { options: { period: 'D7' } });
      assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
      assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
    }
  );

  it('produce export using period format and Teams unique device type output in json',
    async () => {
      const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
          return `Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
          2019-08-28,0,0,0,0,0,0,7
          `;
        }

        throw 'Invalid request';
      });

      await mockCommand.action(logger, { options: { period: 'D7', output: 'json' } });
      assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
      assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
    }
  );

  it('produce export using period format and Teams unique users output in txt',
    async () => {
      const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
          return `
          Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
          2019-08-28,0,0,0,0,0,0,7
          `;
        }

        throw 'Invalid request';
      });

      await mockCommand.action(logger, { options: { period: 'D7', output: 'text' } });
      assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
      assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
    }
  );

  it('produce export using period format and Teams unique users output in json',
    async () => {
      const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
          return `
          Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period
          2019-08-28,0,0,0,0,0,0,7
          `;
        }

        throw 'Invalid request';
      });

      await mockCommand.action(logger, { options: { period: 'D7' } });
      assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
      assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
    }
  );

  it('produce export using period format and Teams output in json',
    async () => {
      const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')`) {
          return `Report Refresh Date,Web,Windows Phone,Android Phone,iOS,Mac,Windows,Report Period\n2019-08-28,0,0,0,0,0,0,7`;
        }

        throw 'Invalid request';
      });

      await mockCommand.action(logger, { options: { debug: true, period: 'D7', output: 'json' } });
      assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/MockEndPoint(period='D7')");
      assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
    }
  );

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(mockCommand.action(logger, { options: { period: 'D7' } } as any), new CommandError('An error has occurred'));
  });
});
