import assert from 'assert';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './user-hibp.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.USER_HIBP, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_HIBP);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userName and apiKey is not specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: 'invalid', apiKey: 'key' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userName and apiKey is specified', async () => {
    const actual = await command.validate({ options: { userName: "account-exists@hibp-integration-tests.com", apiKey: "2975xc539c304xf797f665x43f8x557x" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if domain is specified', async () => {
    const actual = await command.validate({ options: { userName: "account-exists@hibp-integration-tests.com", apiKey: "2975xc539c304xf797f665x43f8x557x", domain: "domain.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('checks user is pwned using userName', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter('account-exists@hibp-integration-tests.com')}`) {
        return [{ "Name": "Adobe" }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: 'account-exists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } });
    assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
  });

  it('checks user is pwned using userName (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter('account-exists@hibp-integration-tests.com')}`) {
        // this is the actual truncated response as the API would return
        return [{ "Name": "Adobe" }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, userName: 'account-exists@hibp-integration-tests.com', apiKey: '2975xc539c304xf797f665x43f8x557x' } });
    assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
  });

  it('checks user is pwned using userName and multiple breaches', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter('account-exists@hibp-integration-tests.com')}`) {
        // this is the actual truncated response as the API would return
        return [{ "Name": "Adobe" }, { "Name": "Gawker" }, { "Name": "Stratfor" }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: 'account-exists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } });
    assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }, { "Name": "Gawker" }, { "Name": "Stratfor" }]));
  });

  it('checks user is pwned using userName and domain', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter('account-exists@hibp-integration-tests.com')}?domain=adobe.com`) {
        // this is the actual truncated response as the API would return
        return [{ "Name": "Adobe" }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: 'account-exists@hibp-integration-tests.com', domain: "adobe.com", apiKey: "2975xc539c304xf797f665x43f8x557x" } });
    assert(loggerLogSpy.calledWith([{ "Name": "Adobe" }]));
  });

  it('checks user is pwned using userName and domain with a domain that does not exists',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter('account-exists@hibp-integration-tests.com')}?domain=adobe.xxx`) {
          // this is the actual truncated response as the API would return
          throw {
            "response": {
              "status": 404
            }
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, userName: 'account-exists@hibp-integration-tests.com', domain: "adobe.xxx", apiKey: "2975xc539c304xf797f665x43f8x557x" } });
      assert(loggerLogSpy.calledWith("No pwnage found"));
    }
  );

  it('correctly handles no pwnage found (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({
      "response": {
        "status": 404
      }
    });

    await command.action(logger, { options: { debug: true, userName: 'account-notexists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } });
    assert(loggerLogSpy.calledWith("No pwnage found"));
  });

  it('correctly handles no pwnage found (verbose)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({
      "response": {
        "status": 404
      }
    });

    await command.action(logger, { options: { verbose: true, userName: 'account-notexists@hibp-integration-tests.com', apiKey: "2975xc539c304xf797f665x43f8x557x" } });
    assert(loggerLogSpy.calledWith("No pwnage found"));
  });

  it('correctly handles unauthorized request', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error("Access denied due to improperly formed hibp-api-key."));

    await assert.rejects(command.action(logger, { options: { userName: 'account-notexists@hibp-integration-tests.com' } } as any),
      new CommandError("Access denied due to improperly formed hibp-api-key."));
  });

  it('fails validation if the userName is not a valid UPN.', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        userName: "no-an-email"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
