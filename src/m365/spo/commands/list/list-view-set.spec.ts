import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './list-view-set.js';

describe(commands.LIST_VIEW_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const listId = '0cd891ef-afce-4e55-b836-fce03286cccf';
  const listTitle = 'List 1';
  const listUrl = '/lists/List 1';
  const viewId = 'cc27a922-8224-4296-90a5-ebbc54da2e81';
  const viewTitle = 'All items';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(spo, 'getRequestDigest').mockClear().mockImplementation().resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: webUrl
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.patch
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_VIEW_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('allows unknown options', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', listTitle: listTitle, title: viewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: 'invalid', title: viewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: 'List 1', id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id and listId specified as valid GUIDs',
    async () => {
      const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, id: viewId } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('ignores global options when creating request data', async () => {
    const patchRequest: jest.Mock = jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events' })) {
          return;
        }
      }

      return 'Invalid request';
    });

    await command.action(logger, { options: { verbose: false, output: "text", webUrl: webUrl, listTitle: listTitle, title: viewTitle, Title: 'All events' } });
    assert.deepEqual(patchRequest.mock.lastCall[0].data, { Title: 'All events' });
  });

  it('updates the Title of the list view specified using its name',
    async () => {
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0 &&
            opts.headers['X-RequestDigest'] &&
            JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events' })) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, title: viewTitle, Title: 'All events' } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('updates the Title and CustomFormatter of the list view specified using its ID',
    async () => {
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views/GetById('${formatting.encodeQueryParameter(viewId)}')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0 &&
            opts.headers['X-RequestDigest'] &&
            JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events', CustomFormatter: 'abc' })) {
            return;
          }
        }

        return Promise.reject('Invalid request');
      });

      await command.action(logger, { options: { debug: true, webUrl: webUrl, listId: listId, id: viewId, Title: 'All events', CustomFormatter: 'abc' } });
    }
  );

  it('updates the Title and CustomFormatter of the list view specified using its Url',
    async () => {
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

        if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views/GetById('${formatting.encodeQueryParameter(viewId)}')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0 &&
            opts.headers['X-RequestDigest'] &&
            JSON.stringify(opts.data) === JSON.stringify({ Title: 'All events', CustomFormatter: 'abc' })) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: webUrl, listUrl: listUrl, id: viewId, Title: 'All events', CustomFormatter: 'abc' } });
    }
  );

  it('correctly handles error when updating existing list view', async () => {
    const errorMessage = 'request rejected';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: errorMessage
          }
        }
      }
    };
    jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        title: viewTitle
      }
    }), new CommandError(errorMessage));
  });
});
