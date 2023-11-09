import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './user-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.USER_GET, () => {
  let cli: Cli;
  let log: any[];
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
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

  it('retrieves user by id with output option json', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetById') > -1) {
        return {
          "value": [{
            "Id": 6,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            "Title": "John Doe",
            "PrincipalType": 1,
            "Email": "john.deo@mytenant.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
            "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    });
    assert(loggerLogSpy.calledWith({
      value: [{
        Id: 6,
        IsHiddenInUI: false,
        LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
        Title: "John Doe",
        PrincipalType: 1,
        Email: "john.deo@mytenant.onmicrosoft.com",
        Expiration: "",
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
        UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
      }]
    }));
  });

  it('retrieves user by email with output option json', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetByEmail') > -1) {
        return {
          "value": [{
            "Id": 6,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            "Title": "John Doe",
            "PrincipalType": 1,
            "Email": "john.deo@mytenant.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
            "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        email: "john.deo@mytenant.onmicrosoft.com"
      }
    });
    assert(loggerLogSpy.calledWith({
      value: [{
        Id: 6,
        IsHiddenInUI: false,
        LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
        Title: "John Doe",
        PrincipalType: 1,
        Email: "john.deo@mytenant.onmicrosoft.com",
        Expiration: "",
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
        UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
      }]
    }));
  });

  it('retrieves user by loginName with output option json', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/siteusers/GetByLoginName') > -1) {
        return {
          "value": [{
            "Id": 6,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            "Title": "John Doe",
            "PrincipalType": 1,
            "Email": "john.deo@mytenant.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
            "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    });
    assert(loggerLogSpy.calledWith({
      value: [{
        Id: 6,
        IsHiddenInUI: false,
        LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
        Title: "John Doe",
        PrincipalType: 1,
        Email: "john.deo@mytenant.onmicrosoft.com",
        Expiration: "",
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
        UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
      }]
    }));
  });


  it('retrieves current logged-in user', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/currentuser') {
        return {
          "value": [{
            "Id": 6,
            "IsHiddenInUI": false,
            "LoginName": "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
            "Title": "John Doe",
            "PrincipalType": 1,
            "Email": "john.deo@mytenant.onmicrosoft.com",
            "Expiration": "",
            "IsEmailAuthenticationGuestUser": false,
            "IsShareByEmailGuestUser": false,
            "IsSiteAdmin": false,
            "UserId": { "NameId": "10010001b0c19a2", "NameIdIssuer": "urn:federation:microsoftonline" },
            "UserPrincipalName": "john.deo@mytenant.onmicrosoft.com"
          }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith({
      value: [{
        Id: 6,
        IsHiddenInUI: false,
        LoginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com",
        Title: "John Doe",
        PrincipalType: 1,
        Email: "john.deo@mytenant.onmicrosoft.com",
        Expiration: "",
        IsEmailAuthenticationGuestUser: false,
        IsShareByEmailGuestUser: false,
        IsSiteAdmin: false,
        UserId: { NameId: "10010001b0c19a2", NameIdIssuer: "urn:federation:microsoftonline" },
        UserPrincipalName: "john.deo@mytenant.onmicrosoft.com"
      }]
    }));
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', id: 1 } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );


  it('fails validation if id, email and loginName options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com", loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if id and email both are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, email: "jonh.deo@mytenant.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if id and loginName options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1, loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if email and loginName options are passed (multiple options)',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com", loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if specified id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid and id is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and email is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', email: "jonh.deo@mytenant.com" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the url is valid and loginName is passed',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com" } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if the url is valid and no other options are provided',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
}); 
