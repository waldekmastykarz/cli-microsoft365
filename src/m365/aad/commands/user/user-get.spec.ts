import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './user-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.USER_GET, () => {
  const userId = "68be84bf-a585-4776-80b3-30aa5207aa21";
  const userName = "AarifS@contoso.onmicrosoft.com";
  const resultValue = { "id": "68be84bf-a585-4776-80b3-30aa5207aa21", "businessPhones": ["+1 425 555 0100"], "displayName": "Aarif Sherzai", "givenName": "Aarif", "jobTitle": "Administrative", "mail": null, "mobilePhone": "+1 425 555 0100", "officeLocation": null, "preferredLanguage": null, "surname": "Sherzai", "userPrincipalName": "AarifS@contoso.onmicrosoft.com" };

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      accessToken.getUserIdFromAccessToken,
      accessToken.getUserNameFromAccessToken,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves user using id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) {
        return { value: [resultValue] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: userId } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using @userid token', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) {
        return { value: [resultValue] };
      }

      throw 'Invalid request';
    });

    jest.spyOn(accessToken, 'getUserIdFromAccessToken').mockClear().mockImplementation(() => { return userId; });

    await command.action(logger, { options: { id: '@meid' } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using id (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) {
        return { value: [resultValue] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: userId } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using user name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) {
        return { value: [resultValue] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using user name and with their direct manager',
    async () => {
      const resultValueWithManger: any = { ...resultValue };
      resultValueWithManger.manager = {
        "displayName": "John Doe",
        "userPrincipalName": "john.doe@contoso.onmicrosoft.com",
        "id": "eb77fbcf-6fe8-458b-985d-1747284793bc",
        "mail": "john.doe@contoso.onmicrosoft.com"
      };
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'&$expand=manager($select=displayName,userPrincipalName,id,mail)`) {
          return { value: [resultValueWithManger] };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { userName: userName, withManager: true } });
      assert(loggerLogSpy.calledWith(resultValueWithManger));
    }
  );

  it('retrieves user using @meusername token', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) {
        return { value: [resultValue] };
      }

      throw 'Invalid request';
    });

    jest.spyOn(accessToken, 'getUserNameFromAccessToken').mockClear().mockImplementation(() => { return userName; });

    await command.action(logger, { options: { userName: '@meusername' } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves user using email', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(userName)}'`) {
        return { value: [resultValue] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { email: userName } });
    assert(loggerLogSpy.calledWith(resultValue));
  });

  it('retrieves only the specified properties', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'&$select=id,mail`) {
        return { value: [{ "id": "userId", "mail": null }] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, properties: 'id,mail' } });
    assert(loggerLogSpy.calledWith({ "id": "userId", "mail": null }));
  });

  it('correctly handles user not found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({
      "error": {
        "code": "Request_ResourceNotFound",
        "message": "Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.",
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
          "date": "2018-04-24T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } } as any),
      new CommandError(`Resource '68be84bf-a585-4776-80b3-30aa5207aa22' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('fails to get user when user with provided id does not exists',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=id eq '${userId}'`) {
          return { value: [] };
        }

        throw `The specified user with id ${userId} does not exist`;
      });

      await assert.rejects(command.action(logger, { options: { id: userId } }),
        new CommandError(`The specified user with id ${userId} does not exist`));
    }
  );

  it('fails to get user when user with provided user name does not exists',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) {
          return { value: [] };
        }

        throw `The specified user with user name ${userName} does not exist`;
      });

      await assert.rejects(command.action(logger, { options: { userName: userName } }),
        new CommandError(`The specified user with user name ${userName} does not exist`));
    }
  );

  it('fails to get user when user with provided email does not exists',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(userName)}'`) {
          return { value: [] };
        }

        throw `The specified user with email ${userName} does not exist`;
      });

      await assert.rejects(command.action(logger, { options: { email: userName } }),
        new CommandError(`The specified user with email ${userName} does not exist`));
    }
  );

  it('handles error when multiple users with the specified email found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(userName)}'`) {
          return {
            value: [
              resultValue,
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', userPrincipalName: 'DebraB@contoso.onmicrosoft.com' }
            ]
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          email: userName
        }
      }), new CommandError("Multiple users with email AarifS@contoso.onmicrosoft.com found. Found: 68be84bf-a585-4776-80b3-30aa5207aa21, 9b1b1e42-794b-4c71-93ac-5ed92488b67f."));
    }
  );

  it('handles selecting single result when multiple users with the specified email found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if ((opts.url as string).indexOf('https://graph.microsoft.com/v1.0/users?$filter') > -1) {
          return {
            value: [
              resultValue,
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', userPrincipalName: 'DebraB@contoso.onmicrosoft.com' }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves(resultValue);

      await command.action(logger, { options: { email: userName } });
      assert(loggerLogSpy.calledWith(resultValue));
    }
  );

  it('fails validation if id or email or userName options are not passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: {} }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if id, email, and userName options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com", userName: "i:0#.f|membership|john.doe@contoso.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both id and email options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", email: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both id and userName options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { id: "1caf7dcd-7e83-4c3a-94f7-932a1299c844", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both email and userName options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { email: "jonh.deo@contoso.com", userName: "john.doe@contoso.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when userName has an invalid value', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userName is specified', async () => {
    const actual = await command.validate({ options: { userName: 'john.doe@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the email is specified', async () => {
    const actual = await command.validate({ options: { email: 'john.doe@contoso.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
