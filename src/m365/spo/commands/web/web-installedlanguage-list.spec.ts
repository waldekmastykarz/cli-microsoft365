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
import command from './web-installedlanguage-list.js';

describe(commands.WEB_INSTALLEDLANGUAGE_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
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
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEB_INSTALLEDLANGUAGE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['DisplayName', 'LanguageTag', 'Lcid']);
  });

  it('retrieves all web installed languages', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/RegionalSettings/InstalledLanguages') > -1) {
        return {
          "Items": [{
            "DisplayName": "German",
            "LanguageTag": "de-DE",
            "Lcid": 1031
          },
          {
            "DisplayName": "French",
            "LanguageTag": "fr-FR",
            "Lcid": 1036
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith([{
      "DisplayName": "German",
      "LanguageTag": "de-DE",
      "Lcid": 1031
    },
    {
      "DisplayName": "French",
      "LanguageTag": "fr-FR",
      "Lcid": 1036
    }]));
  });

  it('command correctly handles web list installed languages reject request',
    async () => {
      const err = 'Invalid request';
      jest.spyOn(request, 'get').mockClear().mockImplementation((opts) => {
        if ((opts.url as string).indexOf('/_api/web/RegionalSettings/InstalledLanguages') > -1) {
          throw err;
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com'
        }
      } as any), new CommandError(err));
    }
  );

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
      const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
