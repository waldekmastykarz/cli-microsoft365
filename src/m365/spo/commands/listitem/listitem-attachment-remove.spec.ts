import assert from 'assert';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './listitem-attachment-remove.js';

describe(commands.LISTITEM_ATTACHMENT_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = '4fc5ba1e-18b7-49e0-81fe-54515cc2eede';
  const listTitle = 'Demo List';
  const listUrl = '/sites/project-x/Lists/DemoList';
  const listItemId = 147;
  const fileName = 'File1.jpg';

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let promptOptions: any;

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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ATTACHMENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', listTitle: listTitle, listItemId: listItemId, fileName: fileName } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the webUrl option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: listTitle, listItemId: listItemId, fileName: fileName } }, commandInfo);
      assert(actual);
    }
  );

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', listItemId: listItemId, fileName: fileName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: listId, listItemId: listItemId, fileName: fileName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the specified listItemId is not a number',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: listTitle, listItemId: 'a', fileName: fileName } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('prompts before removing attachment from list item when confirmation argument not passed (listId)',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listId: listId,
          listItemId: listItemId,
          fileName: fileName
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing attachment from list item when force argument not passed (listTitle)',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listTitle: listTitle,
          listItemId: listItemId,
          fileName: fileName
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing attachment from list item when prompt not confirmed',
    async () => {
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          listId: listId,
          listItemId: listItemId,
          fileName: fileName
        }
      });
    }
  );

  it('removes attachment from list item when listId option is passed and prompt confirmed',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'4fc5ba1e-18b7-49e0-81fe-54515cc2eede')/items(147)/AttachmentFiles('File1.jpg')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
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
          listItemId: listItemId,
          fileName: fileName
        }
      });
      assert(postStub.called);
      assert(loggerLogSpy.notCalled);
    }
  );

  it('removes attachment from list item when listTitle option is passed and prompt confirmed',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/items(147)/AttachmentFiles('File1.jpg')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
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
          listTitle: listTitle,
          listItemId: listItemId,
          fileName: fileName
        }
      });
      assert(postStub.called);
      assert(loggerLogSpy.notCalled);
    }
  );

  it('removes attachment from list item when listUrl option is passed and prompt confirmed',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('%2Fsites%2Fproject-x%2FLists%2FDemoList')/items(147)/AttachmentFiles('File1.jpg')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
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
          listItemId: listItemId,
          fileName: fileName
        }
      });
      assert(postStub.called);
      assert(loggerLogSpy.notCalled);
    }
  );

  it('command correctly handles list get reject request', async () => {
    const err = 'Invalid request';
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/items(147)/AttachmentFiles('File1.jpg')`) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        listTitle: listTitle,
        listItemId: listItemId,
        fileName: fileName,
        force: true
      }
    }), new CommandError(err));
  });

  it('uses correct API url when listTitle option is passed', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('Demo%20List')/items(147)/AttachmentFiles('File1.jpg')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        listTitle: listTitle,
        listItemId: listItemId,
        fileName: fileName,
        force: true
      }
    });
    assert(postStub.called);
    assert(loggerLogSpy.notCalled);
  });
});