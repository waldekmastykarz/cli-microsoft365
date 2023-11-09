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
import command from './list-list.js';

describe(commands.LIST_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  const listResponse = {
    value: [{
      AllowContentTypes: true,
      BaseTemplate: 109,
      BaseType: 1,
      ContentTypesEnabled: false,
      CrawlNonDefaultViews: false,
      Created: null,
      CurrentChangeToken: null,
      CustomActionElements: null,
      DefaultContentApprovalWorkflowId: "00000000-0000-0000-0000-000000000000",
      DefaultItemOpenUseListSetting: false,
      Description: "",
      Direction: "none",
      DocumentTemplateUrl: null,
      DraftVersionVisibility: 0,
      EnableAttachments: false,
      EnableFolderCreation: true,
      EnableMinorVersions: false,
      EnableModeration: false,
      EnableVersioning: false,
      EntityTypeName: "Documents",
      ExemptFromBlockDownloadOfNonViewableFiles: false,
      FileSavePostProcessingEnabled: false,
      ForceCheckout: false,
      HasExternalDataSource: false,
      Hidden: false,
      Id: "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
      ImagePath: null,
      ImageUrl: null,
      IrmEnabled: false,
      IrmExpire: false,
      IrmReject: false,
      IsApplicationList: false,
      IsCatalog: false,
      IsPrivate: false,
      ItemCount: 69,
      LastItemDeletedDate: null,
      LastItemModifiedDate: null,
      LastItemUserModifiedDate: null,
      ListExperienceOptions: 0,
      ListItemEntityTypeFullName: null,
      MajorVersionLimit: 0,
      MajorWithMinorVersionsLimit: 0,
      MultipleDataList: false,
      NoCrawl: false,
      ParentWebPath: null,
      ParentWebUrl: null,
      ParserDisabled: false,
      ServerTemplateCanCreateFolders: true,
      TemplateFeatureId: null,
      Title: "Documents",
      RootFolder: { ServerRelativeUrl: "Documents" }
    }]
  };

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
    assert.strictEqual(command.name, commands.LIST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'Url', 'Id']);
  });

  it('retrieves all lists', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,*') {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith(listResponse.value));
  });

  it('retrieves all lists when properties provided', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate,ParentWebUrl') {
        return {
          value: [{
            BaseTemplate: 100,
            ParentWebUrl: "/",
            RootFolder: {
              ServerRelativeUrl: "/Lists/test"
            },
            Url: "/Lists/test"
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        properties: 'BaseTemplate,ParentWebUrl,RootFolder/ServerRelativeUrl'
      }
    });
    assert(loggerLogSpy.calledWith([{
      BaseTemplate: 100,
      ParentWebUrl: "/",
      RootFolder: {
        ServerRelativeUrl: "/Lists/test"
      },
      Url: "/Lists/test"
    }]));
  });

  it('retrieves the lists based on the provided filter', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,*&$filter=BaseTemplate eq 100') {
        return listResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        filter: 'BaseTemplate eq 100'
      }
    });
    assert(loggerLogSpy.calledWith(listResponse.value));
  });

  it('command correctly handles list list reject request', async () => {
    const err = 'Invalid request';
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/lists?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,*') {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
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

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
      assert(actual);
    }
  );
});
