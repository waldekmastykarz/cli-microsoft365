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
import commands from '../../commands.js';
import command from './list-webhook-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LIST_WEBHOOK_REMOVE, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
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
    requests = [];
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing webhook from list when confirmation argument not passed (list title)',
    async () => {
      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing webhook from list when confirmation argument not passed (list url)',
    async () => {
      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing list when confirmation argument not passed (list id)',
    async () => {
      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing list when prompt not confirmed', async () => {
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
    assert(requests.length === 0);
  });

  it('removes the list (retrieved by Title) when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
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

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by Title) webhook when prompt confirmed',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
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

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by id) webhook when prompt confirmed',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
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

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: 'dfddade1-4729-428d-881e-7fedf3cae50d', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by id) webhook when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
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

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: 'dfddade1-4729-428d-881e-7fedf3cae50d', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by id) webhook when prompt confirmed in options',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: 'dfddade1-4729-428d-881e-7fedf3cae50d', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81', force: true } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by url) webhook when confirmed in options',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/lists/Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81', force: true } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by url) webhook when prompt confirmed',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
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

      await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/lists/Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('%2Fsites%2Fninja%2Flists%2FDocuments')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('removes the list (retrieved by url) webhook when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);

        if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
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

      await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/lists/Documents', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
  );

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        force: true
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if webhook id option is not passed', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', id: 'cc27a922-8224-4296-90a5-ebbc54da2e85', listId: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, commandInfo);
      assert(actual);
    }
  );

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listid option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the listid option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', id: 'cc27a922-8224-4296-90a5-ebbc54da2e81' } }, commandInfo);
    assert(actual);
  });
});
