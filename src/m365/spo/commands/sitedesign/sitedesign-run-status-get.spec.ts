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
import command from './sitedesign-run-status-get.js';

describe(commands.SITEDESIGN_RUN_STATUS_GET, () => {
  let log: string[];
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
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_RUN_STATUS_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ActionTitle', 'SiteScriptTitle', 'OutcomeText']);
  });

  it('gets information about site designs applied to the specified site',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`) > -1) {
          return {
            "value": [
              { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
              { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } });
      assert(loggerLogSpy.calledWith([
        { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
        { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
      ]));
    }
  );

  it('outputs all information in JSON output mode', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`) > -1) {
        return {
          "value": [
            { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
            { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8', output: 'json' } });
    assert(loggerLogSpy.calledWith([
      { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
      { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
    ]));
  });

  it('correctly handles error when the specified runId doesn\'t point to a valid run',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({ error: { 'odata.error': { message: { value: 'Value does not fall within the expected range' } } } });

      await assert.rejects(command.action(logger, {
        options: {
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8'
        }
      } as any), new CommandError('Value does not fall within the expected range'));
    }
  );

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if runId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', runId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and runId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', runId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
