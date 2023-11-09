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
import { ClientSidePage } from './clientsidepages.js';
import command from './page-control-get.js';
import { mockControlGetData, mockControlGetDataEmptyColumn, mockControlGetDataEmptyColumnOutput, mockControlGetDataOutput, mockControlGetDataWithoutAnId, mockControlGetDataWithText, mockControlGetDataWithTextOutput, mockControlGetDataWithUnknownType, mockControlGetDataWithUnknownTypeOutput } from './page-control-get.mock.js';

describe(commands.PAGE_CONTROL_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: jest.SpyInstance;

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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      ClientSidePage.fromHtml
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_CONTROL_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the control on a modern page', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlGetData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } });
    assert(loggerLogSpy.calledWith(mockControlGetDataOutput));
  });

  it('gets information about the control on a modern page (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
          return mockControlGetData;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } });
      assert(loggerLogSpy.calledWith(mockControlGetDataOutput));
    }
  );

  it('gets information about the control on a modern page when the specified page name doesn\'t contain extension',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
          return mockControlGetData;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } });
      assert(loggerLogSpy.calledWith(mockControlGetDataOutput));
    }
  );

  it('handles empty columns', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlGetDataEmptyColumn;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '88f7b5b2-83a8-45d1-bc61-c11425f233e3' } });
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(mockControlGetDataEmptyColumnOutput));
  });

  it('handles text controls', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlGetDataWithText;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '1212fc8d-dd6b-408a-8d5d-9f1cc787efbb' } });
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(mockControlGetDataWithTextOutput));
  });

  it('handles unknown types', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlGetDataWithUnknownType;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: 'af92a21f-a0ec-4668-ba2c-951a2b5d6f94' } });
    assert(loggerLogSpy.calledWith(mockControlGetDataWithUnknownTypeOutput));
  });


  it('correctly handles control not found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlGetData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles control not found (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
        return mockControlGetData;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } });
    assert(loggerLogToStderrSpy.calledWith('Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx'));
  });

  it('correctly handles control not found and no ID was specified',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
          return mockControlGetDataWithoutAnId;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } });
      assert(loggerLogToStderrSpy.calledWith('Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx'));
    }
  );

  it('correctly handles control not found on a page when CanvasContent1 is not defined',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SitePages/Pages/GetByUrl('sitepages/home.aspx')`) > -1) {
          return { CanvasContent1: null };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx', id: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e6' } });
      assert(loggerLogToStderrSpy.calledWith('Control with ID 3ede60d3-dc2c-438b-b5bf-cc40bb2351e6 not found on page home.aspx'));
    }
  );

  it('correctly handles page not found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      throw {
        error: {
          "odata.error": {
            "code": "-2130575338, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "The file /sites/team-a/SitePages/home1.aspx does not exist."
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } } as any),
      new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.'));
  });

  it('correctly handles OData error when retrieving pages', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', pageName: 'home.aspx' } } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the specified id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc', pageName: 'home.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when the webUrl is a valid SharePoint URL and name and id are specified',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', pageName: 'home.aspx', id: 'ede2ee65-157d-4523-b4ed-87b9b64374a6' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
