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
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './list-view-remove.js';

describe(commands.LIST_VIEW_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/ninja';
  const listId = '0cd891ef-afce-4e55-b836-fce03286cccf';
  const listTitle = 'Documents';
  const listUrl = '/sites/ninja/Shared Documents';
  const viewId = 'cc27a922-8224-4296-90a5-ebbc54da2e81';
  const viewTitle = 'MyView';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

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
    assert.strictEqual(command.name, commands.LIST_VIEW_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', id: viewId, listId: listId } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: '12345', id: viewId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle, title: viewTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified view from list by id and listTitle when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listTitle: listTitle,
          id: viewId
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing the specified view from list by title and listId when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listId: listId,
          title: listTitle
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing the specified view from list by title and listUrl when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listUrl: listUrl,
          title: listTitle
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing view from list when prompt not confirmed', async () => {
    const postSpy = jest.spyOn(request, 'post').mockClear();

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        id: viewId
      }
    });

    assert(postSpy.notCalled);
  });

  it('removes view from the list using id and listUrl when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
        if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views(guid'${formatting.encodeQueryParameter(viewId)}')`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: webUrl,
          listUrl: listUrl,
          id: viewId
        }
      });
    }
  );

  it('removes view from the list using id and listId when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views(guid'${formatting.encodeQueryParameter(viewId)}')`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: webUrl,
          listId: listId,
          id: viewId,
          force: true
        }
      });
    }
  );

  it('removes view from the list using id and listTitle when prompt confirmed',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views(guid'${formatting.encodeQueryParameter(viewId)}')`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listTitle: listTitle,
          id: viewId,
          force: true
        }
      });
    }
  );

  it('removes view from the list using title and listUrl when prompt confirmed',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
        if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listUrl: listUrl,
          title: viewTitle,
          force: true
        }
      });
    }
  );

  it('removes view from the list using title and listId when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          debug: true,
          webUrl: webUrl,
          listId: listId,
          title: viewTitle,
          force: true
        }
      });
    }
  );

  it('removes view from the list using title and listTitle when prompt confirmed',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views/GetByTitle('${formatting.encodeQueryParameter(viewTitle)}')`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listTitle: listTitle,
          title: viewTitle,
          force: true
        }
      });
    }
  );

  it('correctly handles error when removing view from the list', async () => {
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
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        listTitle: listTitle,
        title: viewTitle,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
