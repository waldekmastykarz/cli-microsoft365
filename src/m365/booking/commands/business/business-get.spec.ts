import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './business-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.BUSINESS_GET, () => {
  let cli: Cli;
  const validId = 'mail@contoso.onmicrosoft.com';
  const validName = 'Valid Business';

  const businessResponse = {
    'id': validId,
    'displayName': validName,
    'businessType': 'Other',
    'phone': '',
    'email': 'user@contoso.onmicrosoft.com',
    'webSiteUrl': '',
    'defaultCurrencyIso': 'USD'
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  beforeAll(() => {
    cli = Cli.getInstance();
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      Cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.BUSINESS_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the text output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'businessType', 'phone', 'email', 'defaultCurrencyIso']);
  });

  it('gets business by id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return businessResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validId } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('gets business by title', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [businessResponse] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return businessResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: validName } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('fails when multiple businesses found with same name', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [businessResponse, businessResponse] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError("Multiple businesses with name 'Valid Business' found. Found: mail@contoso.onmicrosoft.com."));
  });

  it('handles selecting single result when multiple businesses with the specified name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
          return Promise.resolve({ value: [businessResponse, businessResponse] });
        }

        if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
          return Promise.resolve(businessResponse);
        }

        return Promise.reject('Invalid request');
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves(businessResponse);

      await command.action(logger, { options: { name: validName } });
      assert(loggerLogSpy.calledWith(businessResponse));
    }
  );

  it('fails when no business found with name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
  });

  it('fails when no business found with name because of an empty displayName',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
          return { value: [{ 'displayName': null }] };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
    }
  );

  it('correctly handles random API error', async () => {
    jestUtil.restore(request.get);
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
