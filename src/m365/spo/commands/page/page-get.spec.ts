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
import command from './page-get.js';
import { classicPage, controlsMock, pageListItemMock, sectionMock } from './page-get.mock.js';

describe(commands.PAGE_GET, () => {
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
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['commentsDisabled', 'numSections', 'numControls', 'title', 'layoutType']);
  });

  it('gets information about a modern page including all returned properties',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/SitePages/home.aspx')`) > -1) {
          return pageListItemMock;
        }

        if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
          return controlsMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', output: 'json' } } as any);
      assert.strictEqual(loggerLogSpy.mock.lastCall[0].numControls, sectionMock.numControls);
      assert.strictEqual(loggerLogSpy.mock.lastCall[0].numSections, sectionMock.numSections);
    }
  );

  it('gets information about a modern page', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return controlsMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any);
    assert(loggerLogSpy.calledWith({
      ...pageListItemMock,
      canvasContentJson: controlsMock.CanvasContent1,
      ...sectionMock
    }));
  });

  it('gets information about a modern page on root of tenant', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/SitePages/home.aspx')`) > -1) {
        return pageListItemMock;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return controlsMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx', output: 'json' } } as any);
    assert(loggerLogSpy.calledWith({
      ...pageListItemMock,
      canvasContentJson: controlsMock.CanvasContent1,
      ...sectionMock
    }));
  });

  it('gets information about a modern page when the specified page name doesn\'t contain extension',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/SitePages/home.aspx')`) > -1) {
          return pageListItemMock;
        }

        if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
          return controlsMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', output: 'json' } } as any);
      assert(loggerLogSpy.calledWith({
        ...pageListItemMock,
        canvasContentJson: controlsMock.CanvasContent1,
        ...sectionMock
      }));
    }
  );

  it('check if section and control HTML parsing gets skipped for metadata only mode',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/SitePages/home.aspx')`) > -1) {
          return pageListItemMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', metadataOnly: true, output: 'json' } });
      assert(loggerLogSpy.calledWith(pageListItemMock));
    }
  );

  it('shows error when the specified page is a classic page', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/SitePages/home.aspx')`) > -1) {
        return classicPage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any),
      new CommandError('Page home.aspx is not a modern page.'));
  });

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

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any),
      new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.'));
  });

  it('correctly handles OData error when retrieving pages', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying metadataOnly flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--metadataOnly') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', name: 'home.aspx' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when the webUrl is a valid SharePoint URL and name is specified',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
