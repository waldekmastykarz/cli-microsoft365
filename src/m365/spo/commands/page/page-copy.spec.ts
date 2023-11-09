import assert from 'assert';
import chalk from 'chalk';
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
import command from './page-copy.js';
import { copyMock } from './page-copy.mock.js';

describe(commands.PAGE_COPY, () => {
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
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_COPY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'PageLayoutType', 'Title', 'Url']);
  });

  it('create a page copy', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy (DEBUG)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx" } });
    assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
  });

  it('create a page copy and automatically append the aspx extension',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
          return copyMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } });
      assert(loggerLogSpy.calledWith(copyMock));
    }
  );

  it('create a page copy and check if the webUrl is automatically added',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
          return copyMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "home-copy" } });
      assert(loggerLogSpy.calledWith(copyMock));
    }
  );

  it('create a page copy with leading slash in the targetUrl', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and check if correct URL is used when sitepages is already added',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
          return copyMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "sitepages/home-copy" } });
      assert(loggerLogSpy.calledWith(copyMock));
    }
  );

  it('create a page copy and check if correct URL is used when sitepages (with leading slash) is already added',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
          return copyMock;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "/sitepages/home-copy" } });
      assert(loggerLogSpy.calledWith(copyMock));
    }
  );

  it('create a page copy to another site', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/team-b/_api/sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home", targetUrl: "https://contoso.sharepoint.com/sites/team-b/sitepages/home-copy" } });
    assert(loggerLogSpy.calledWith(copyMock));
  });

  it('create a page copy and overwrite the file', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`sitepages/pages/GetByUrl('sitepages/home-copy.aspx')`) > -1) {
        return copyMock;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } });
    assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));
  });

  it('catch any other error in the copy command', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SP.MoveCopyUtil.CopyFileByPath()`) > -1) {
        throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', sourceName: "home.aspx", targetUrl: "home-copy.aspx", overwrite: true } }), new CommandError('An error has occurred'));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', sourceName: 'home.aspx', targetUrl: 'home-copy.aspx' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation when the webUrl is a valid SharePoint URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceName: 'home.aspx', targetUrl: 'home-copy.aspx' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
