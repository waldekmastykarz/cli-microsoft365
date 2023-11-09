import assert from 'assert';
import { telemetry } from '../../telemetry.js';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import request from '../../request.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import SpoCommand from './SpoCommand.js';

class MockCommand extends SpoCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  constructor() {
    super();

    this.options.unshift(
      {
        option: '--url [url]'
      },
      {
        option: '--nonProcessedUrl [nonProcessedUrl]'
      }
    );
  }

  public async commandAction(): Promise<void> {
  }

  public validateUnknownCsomOptionsPublic(options: any, csomObject: string, csomPropertyType: 'get' | 'set'): string | boolean {
    return this.validateUnknownCsomOptions(options, csomObject, csomPropertyType);
  }

  public getNamesOfOptionsWithUrlsPublic(): string[] {
    return this.getNamesOfOptionsWithUrls();
  }
}

describe('SpoCommand', () => {
  let logger: Logger;
  let log: string[];

  beforeAll(() => {
    auth.service.connected = true;
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.post,
      auth.storeConnectionInfo,
      auth.restoreAuth
    ]);
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('correctly reports an error while restoring auth info', async () => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(async () => { throw 'An error has occurred'; });
    const command = new MockCommand();

    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('doesn\'t execute command when error occurred while restoring auth info',
    async () => {
      jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(async () => { throw 'An error has occurred'; });
      const command = new MockCommand();
      const logger: Logger = {
        log: async () => { },
        logRaw: async () => { },
        logToStderr: async () => { }
      };
      const commandCommandActionSpy = jest.spyOn(command, 'commandAction').mockClear();
      await assert.rejects(command.action(logger, { options: {} }));
      assert(commandCommandActionSpy.notCalled);
    }
  );

  it('passes validation of unknown properties when no unknown properties are set',
    async () => {
      const command = new MockCommand();
      assert.strictEqual(command.validateUnknownCsomOptionsPublic({}, 'web', 'set'), true);
    }
  );

  it('passes validation of unknown properties when valid unknown properties specified',
    async () => {
      const command = new MockCommand();
      assert.strictEqual(command.validateUnknownCsomOptionsPublic({ AllowAutomaticASPXPageIndexing: true }, 'web', 'set'), true);
    }
  );

  it('fails validation of unknown properties when invalid unknown property specified',
    async () => {
      const command = new MockCommand();
      assert.notStrictEqual(command.validateUnknownCsomOptionsPublic({ AllowCreateDeclarativeWorkflow: true }, 'web', 'set'), true);
    }
  );

  it('fails validation of unknown properties when unknown property of unsupported type specified',
    async () => {
      const command = new MockCommand();
      assert.notStrictEqual(command.validateUnknownCsomOptionsPublic({ AssociatedMemberGroup: {} }, 'web', 'set'), true);
    }
  );

  it('returns default list of names of options with URLs if no names to exclude defined',
    () => {
      const expected = [
        'appCatalogUrl',
        'actionUrl',
        'imageUrl',
        'libraryUrl',
        'logoUrl',
        'newSiteUrl',
        'NoAccessRedirectUrl',
        'OrgNewsSiteUrl',
        'origin',
        'parentUrl',
        'parentWebUrl',
        'previewImageUrl',
        'siteLogoUrl',
        'siteUrl',
        'StartASiteFormUrl',
        'targetUrl',
        'thumbnailUrl',
        'url',
        'webUrl'
      ];
      const command = new MockCommand();
      const actual = command.getNamesOfOptionsWithUrlsPublic();
      assert.deepStrictEqual(actual, expected);
    }
  );

  it('returns filtered list of names of options with URLs when names to exclude defined',
    () => {
      const expected = [
        'appCatalogUrl',
        'actionUrl',
        'imageUrl',
        'libraryUrl',
        'logoUrl',
        'newSiteUrl',
        'NoAccessRedirectUrl',
        'OrgNewsSiteUrl',
        'origin',
        'parentUrl',
        'parentWebUrl',
        'previewImageUrl',
        'siteLogoUrl',
        'siteUrl',
        'StartASiteFormUrl',
        'targetUrl',
        'thumbnailUrl',
        'webUrl'
      ];
      const command = new MockCommand();
      jest.spyOn(command as any, 'getExcludedOptionsWithUrls').mockClear().mockImplementation(() => ['url']);
      const actual = command.getNamesOfOptionsWithUrlsPublic();
      assert.deepStrictEqual(actual, expected);
    }
  );

  it('resolves server-relative URLs in known options to absolute when SPO URL available',
    async () => {
      const command = new MockCommand();
      auth.service.spoUrl = 'https://contoso.sharepoint.com';
      const options = {
        url: '/'
      };
      await command.processOptions(options);
      assert.strictEqual(options.url, 'https://contoso.sharepoint.com/');
    }
  );

  it('leaves absolute URLs as-is', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: 'https://contoso.sharepoint.com/sites/contoso'
    };
    await command.processOptions(options);
    assert.strictEqual(options.url, 'https://contoso.sharepoint.com/sites/contoso');
  });

  it('leaves site-relative URLs as-is', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      url: 'sites/contoso'
    };
    await command.processOptions(options);
    assert.strictEqual(options.url, 'sites/contoso');
  });

  it('leaves server-relative URLs as-is in unknown options', async () => {
    const command = new MockCommand();
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    const options = {
      nonProcessedUrl: '/'
    };
    await command.processOptions(options);
    assert.strictEqual(options.nonProcessedUrl, '/');
  });

  it('throws error when server-relative URL specified but SPO URL not available',
    async () => {
      const command = new MockCommand();
      const options = {
        url: '/'
      };
      await assert.rejects(command.processOptions(options));
    }
  );

  it('Shows an error when CLI is connected with authType "Secret"',
    async () => {
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(5);

      const mock = new MockCommand();
      await assert.rejects(mock.action(logger, { options: {} }),
        new CommandError('SharePoint does not support authentication using client ID and secret. Please use a different login type to use SharePoint commands.'));
    }
  );
});
