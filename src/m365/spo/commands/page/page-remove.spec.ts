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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './page-remove.js';

describe(commands.PAGE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let promptOptions: any;

  const fakeRestCalls: (pageName?: string) => jest.Mock = (pageName: string = 'page.aspx') => {
    return jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/${pageName}')`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });
  };

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(spo, 'getRequestDigest').mockClear().mockImplementation().resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a modern page without confirm prompt', async () => {
    fakeRestCalls();
    await command.action(logger,
      {
        options: {
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('removes a modern page (debug) without confirm prompt', async () => {
    fakeRestCalls();
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('removes a modern page (debug) without confirm prompt on root of tenant',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sitepages/page.aspx')`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      await command.action(logger,
        {
          options: {
            debug: true,
            name: 'page.aspx',
            webUrl: 'https://contoso.sharepoint.com',
            force: true
          }
        });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('removes a modern page with confirm prompt', async () => {
    fakeRestCalls();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: true };
    });
    await command.action(logger,
      {
        options: {
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('removes a modern page (debug) with confirm prompt', async () => {
    fakeRestCalls();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: true };
    });
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(loggerLogToStderrSpy.called);
  });

  it('should prompt before removing page when confirmation argument not passed',
    async () => {
      fakeRestCalls();
      await command.action(logger,
        {
          options: {
            debug: true,
            name: 'page.aspx',
            webUrl: 'https://contoso.sharepoint.com/sites/team-a'
          }
        });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('should abort page removal when prompt not confirmed', async () => {
    const postCallSpy = fakeRestCalls();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));
    await command.action(logger,
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      });
    assert(postCallSpy.notCalled === true);
  });

  it('automatically appends the .aspx extension', async () => {
    fakeRestCalls();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));
    await command.action(logger,
      {
        options: {
          name: 'page',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when removing modern page', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(() => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));
    await assert.rejects(command.action(logger,
      {
        options: {
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          force: true
        }
      }), new CommandError('An error has occurred'));
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirm', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--force') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({
      options: { name: 'page.aspx', webUrl: 'http://foo' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL',
    async () => {
      const actual = await command.validate({
        options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when name has no extension', async () => {
    const actual = await command.validate({
      options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
