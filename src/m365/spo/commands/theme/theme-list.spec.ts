import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './theme-list.js';

describe(commands.THEME_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let loggerLogToStderrSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name']);
  });

  it('uses correct API url', async () => {
    const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
      }
    });
    assert.strictEqual(postStub.mock.lastCall[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/GetTenantThemingOptions');
    assert.strictEqual(postStub.mock.lastCall[0].headers['accept'], 'application/json;odata=nometadata');
  });

  it('uses correct API url (debug)', async () => {
    const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        return 'Correct Url';
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert.strictEqual(postStub.mock.lastCall[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/GetTenantThemingOptions');
    assert.strictEqual(postStub.mock.lastCall[0].headers['accept'], 'application/json;odata=nometadata');
    assert.strictEqual(loggerLogToStderrSpy.called, true);
  });

  it('retrieves available themes from the tenant store', async () => {
    const themes: any = {
      "themePreviews": [{ "name": "Mint", "themeJson": "{\"isInverted\":false,\"name\":\"Mint\",\"palette\":{\"themePrimary\":\"#43cfbb\",\"themeLighterAlt\":\"#f2fcfa\",\"themeLighter\":\"#ddf6f2\",\"themeLight\":\"#adeae1\",\"themeTertiary\":\"#71dbcb\",\"themeSecondary\":\"#4ad1bd\",\"themeDarkAlt\":\"#32c3ae\",\"themeDark\":\"#248b7b\",\"themeDarker\":\"#1f776a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}" }, { "name": "Mint Inverted", "themeJson": "{\"isInverted\":true,\"name\":\"Mint Inverted\",\"palette\":{\"themePrimary\":\"#43cfbb\",\"themeLighterAlt\":\"#f2fcfa\",\"themeLighter\":\"#ddf6f2\",\"themeLight\":\"#adeae1\",\"themeTertiary\":\"#71dbcb\",\"themeSecondary\":\"#4ad1bd\",\"themeDarkAlt\":\"#32c3ae\",\"themeDark\":\"#248b7b\",\"themeDarker\":\"#1f776a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}" }]
    };
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        return themes;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true } });
    assert(loggerLogSpy.calledWith([{
      "name": "Mint",
      "themeJson": "{\"isInverted\":false,\"name\":\"Mint\",\"palette\":{\"themePrimary\":\"#43cfbb\",\"themeLighterAlt\":\"#f2fcfa\",\"themeLighter\":\"#ddf6f2\",\"themeLight\":\"#adeae1\",\"themeTertiary\":\"#71dbcb\",\"themeSecondary\":\"#4ad1bd\",\"themeDarkAlt\":\"#32c3ae\",\"themeDark\":\"#248b7b\",\"themeDarker\":\"#1f776a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}"
    },
    {
      "name": "Mint Inverted",
      "themeJson": "{\"isInverted\":true,\"name\":\"Mint Inverted\",\"palette\":{\"themePrimary\":\"#43cfbb\",\"themeLighterAlt\":\"#f2fcfa\",\"themeLighter\":\"#ddf6f2\",\"themeLight\":\"#adeae1\",\"themeTertiary\":\"#71dbcb\",\"themeSecondary\":\"#4ad1bd\",\"themeDarkAlt\":\"#32c3ae\",\"themeDark\":\"#248b7b\",\"themeDarker\":\"#1f776a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}"
    }]), 'Invalid request');
  });

  it('retrieves available themes from the tenant store with all properties for JSON output',
    async () => {
      const expected: any = {
        "themePreviews": [{ "name": "Mint", "themeJson": "{\"isInverted\":false,\"name\":\"Mint\",\"palette\":{\"themePrimary\":\"#43cfbb\",\"themeLighterAlt\":\"#f2fcfa\",\"themeLighter\":\"#ddf6f2\",\"themeLight\":\"#adeae1\",\"themeTertiary\":\"#71dbcb\",\"themeSecondary\":\"#4ad1bd\",\"themeDarkAlt\":\"#32c3ae\",\"themeDark\":\"#248b7b\",\"themeDarker\":\"#1f776a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}" }, { "name": "Mint Inverted", "themeJson": "{\"isInverted\":true,\"name\":\"Mint Inverted\",\"palette\":{\"themePrimary\":\"#43cfbb\",\"themeLighterAlt\":\"#f2fcfa\",\"themeLighter\":\"#ddf6f2\",\"themeLight\":\"#adeae1\",\"themeTertiary\":\"#71dbcb\",\"themeSecondary\":\"#4ad1bd\",\"themeDarkAlt\":\"#32c3ae\",\"themeDark\":\"#248b7b\",\"themeDarker\":\"#1f776a\",\"neutralLighterAlt\":\"#f8f8f8\",\"neutralLighter\":\"#f4f4f4\",\"neutralLight\":\"#eaeaea\",\"neutralQuaternaryAlt\":\"#dadada\",\"neutralQuaternary\":\"#d0d0d0\",\"neutralTertiaryAlt\":\"#c8c8c8\",\"neutralTertiary\":\"#a6a6a6\",\"neutralSecondary\":\"#666666\",\"neutralPrimaryAlt\":\"#3c3c3c\",\"neutralPrimary\":\"#333\",\"neutralDark\":\"#212121\",\"black\":\"#1c1c1c\",\"white\":\"#fff\",\"primaryBackground\":\"#fff\",\"primaryText\":\"#333\",\"bodyBackground\":\"#fff\",\"bodyText\":\"#333\",\"disabledBackground\":\"#f4f4f4\",\"disabledText\":\"#c8c8c8\"}}" }]
      };
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
          return expected;
        }
        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, verbose: true, output: 'json' } });
      assert(loggerLogSpy.calledWith(expected.themePreviews), 'Invalid request');
    }
  );

  it('retrieves available themes - handle error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetTenantThemingOptions') > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true } } as any), new CommandError('An error has occurred'));
  });
});
