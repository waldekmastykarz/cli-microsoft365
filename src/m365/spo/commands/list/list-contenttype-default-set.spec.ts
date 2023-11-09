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
import command from './list-contenttype-default-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LIST_CONTENTTYPE_DEFAULT_SET, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_CONTENTTYPE_DEFAULT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures specified visible content type as default. List specified using Title. UniqueContentTypeOrder null',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder`) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('configures specified visible content type as default. List specified using Title. UniqueContentTypeOrder null. Debug',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder`) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogToStderrSpy.called);
      assert(loggerLogSpy.notCalled);
    }
  );

  it('configures specified visible content type as default. List specified using ID. UniqueContentTypeOrder not null',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder`) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ],
            "UniqueContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ]
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('configures specified visible content type as default. List specified using URL. UniqueContentTypeOrder not null',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/RootFolder`) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ],
            "UniqueContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ]
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList(\'%2Fsites%2Fdocuments\')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          listUrl: 'sites/documents',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('configures specified visible content type as default. List specified using ID. UniqueContentTypeOrder not null. Debug',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder`) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ],
            "UniqueContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              },
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              }
            ]
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogToStderrSpy.called);
      assert(loggerLogSpy.notCalled);
    }
  );

  it('configures specified invisible content type as default. List specified using Title. UniqueContentTypeOrder null',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder` &&
          opts.headers &&
          opts.headers['x-http-method'] === 'MERGE' &&
          JSON.stringify(opts.data) === JSON.stringify({
            UniqueContentTypeOrder: [
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              },
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ]
          })) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('configures specified invisible content type as default. List specified using Title. UniqueContentTypeOrder null. Debug',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder` &&
          opts.headers &&
          opts.headers['x-http-method'] === 'MERGE' &&
          JSON.stringify(opts.data) === JSON.stringify({
            UniqueContentTypeOrder: [
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              },
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ]
          })) {
          return;
        }

        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
              },
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
      assert(loggerLogToStderrSpy.called);
      assert(loggerLogSpy.notCalled);
    }
  );

  it(`doesn't configure content type as default if it's already set as default`,
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              },
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
    }
  );

  it(`doesn't configure content type as default if it's already set as default. Debug`,
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
              },
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      });
    }
  );

  it(`fails, if the specified web doesn't exist`, async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
      throw 'Invalid request';
    });
    jest.spyOn(request, 'get').mockClear().mockImplementation(async () => {
      throw 'Request failed with status code 404';
    });

    await assert.rejects(command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    }), new CommandError('Request failed with status code 404'));
  });

  it(`fails, if the list specified by title doesn't exist`, async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
      throw 'Invalid request';
    });
    jest.spyOn(request, 'get').mockClear().mockImplementation(async () => {
      throw 'Request failed with status code 404';
    });

    await assert.rejects(command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    }), new CommandError('Request failed with status code 404'));
  });

  it(`fails, if the specified content type not found in the list`,
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
        throw 'Invalid request';
      });
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
          return {
            "ContentTypeOrder": [
              {
                "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
              }
            ],
            "UniqueContentTypeOrder": null
          };
        }

        if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('My%20List')/ContentTypes?$select=Id`) {
          return {
            value: [
              {
                Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          listTitle: 'My List',
          webUrl: 'https://contoso.sharepoint.com',
          contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
        }
      }), new CommandError('Content type 0x0104001A75DCE30BAC754AA5134C183CF7A92E missing in the list. Add the content type to the list first and try again.'));
    }
  );

  it('fails validation if neither listId nor listTitle nor listUrl are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', contentTypeId: '0x0120' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if all of the list properties are passed', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', listUrl: 'sites/documents', contentTypeId: '0x0120' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the webUrl option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } }, commandInfo);
      assert(actual);
    }
  );

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', contentTypeId: '0x0120' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the listTitle option is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', contentTypeId: '0x0120' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', contentTypeId: '0x0120' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

});
