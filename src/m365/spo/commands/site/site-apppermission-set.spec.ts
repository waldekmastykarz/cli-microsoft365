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
import command from './site-apppermission-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITE_APPPERMISSION_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.patch,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPPERMISSION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "123"
      }
    }, commandInfo);

    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, appId, and appDisplayName options are not specified',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
          permission: "write"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation with a correct URL', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if invalid value specified for permission',
    async () => {
      const actual = await command.validate({
        options: {
          siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
          permission: "Invalid",
          appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails when passing a site that does not exist', async () => {
    const siteError = {
      "error": {
        "code": "itemNotFound",
        "message": "Requested site could not be found",
        "innerError": {
          "date": "2021-03-03T08:58:02",
          "request-id": "4e054f93-0eba-4743-be47-ce36b5f91120",
          "client-request-id": "dbd35b28-0ec3-6496-1279-0e1da3d028fe"
        }
      }
    };
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('non-existing') === -1) {
        return { value: [] };
      }
      throw siteError;
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name-non-existing',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    } as any), new CommandError('Requested site could not be found'));
  });

  it('fails to get Azure AD app when Azure AD app does not exists',
    async () => {
      const getRequestStub = jest.spyOn(request, 'get').mockClear().mockImplementation();
      getRequestStub.onCall(0)
        .callsFake(async (opts) => {
          if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
            return {
              "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
              "displayName": "sitecollection-name",
              "name": "sitecollection-name",
              "createdDateTime": "2021-03-09T20:56:00Z",
              "lastModifiedDateTime": "2021-03-09T20:56:01Z",
              "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
            };
          }
          throw 'Invalid request';
        });

      getRequestStub.onCall(1)
        .callsFake(async (opts) => {
          if ((opts.url as string).indexOf('/permissions') > -1) {
            return { value: [] };
          }
          throw 'The specified app permission does not exist';
        });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
          permission: "write",
          appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
        }
      } as any), new CommandError('The specified app permission does not exist'));
    }
  );

  it('fails when multiple Azure AD apps with same name exists', async () => {
    const getRequestStub = jest.spyOn(request, 'get').mockClear().mockImplementation();
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Multiple app permissions with displayName Foo found: 89ea5c94-7736-4e25-95ad-3fa95f62b66e,cca00169-d38b-462f-a3b4-f3566b162f2d7';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return {
            "value": [
              {
                "id": "aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "Foo",
                      "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
                    }
                  }
                ]
              },
              {
                "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "Foo",
                      "id": "cca00169-d38b-462f-a3b4-f3566b162f2d7"
                    }
                  }
                ]
              }
            ]
          };
        }
        throw 'Multiple app permissions with displayName Foo found: 89ea5c94-7736-4e25-95ad-3fa95f62b66e,cca00169-d38b-462f-a3b4-f3566b162f2d7';
      });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appDisplayName: "Foo"
      }
    } as any), new CommandError('Multiple app permissions with displayName Foo found: 89ea5c94-7736-4e25-95ad-3fa95f62b66e,cca00169-d38b-462f-a3b4-f3566b162f2d7'));
  });

  it('Updates an application permission to the site by appId', async () => {
    const getRequestStub = jest.spyOn(request, 'get').mockClear().mockImplementation();
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return {
            "value": [
              {
                "id": "aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "Foo",
                      "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
                    }
                  }
                ]
              },
              {
                "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                "grantedToIdentities": [
                  {
                    "application": {
                      "displayName": "TeamsBotDemo5",
                      "id": "cca00169-d38b-462f-a3b4-f3566b162f2d"
                    }
                  }
                ]
              }
            ]
          };
        }

        throw 'Invalid request';
      });

    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        output: "json"
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
          }
        }
      ]
    }));
  });

  it('Updates an application permission to the site by appDisplayName',
    async () => {
      const getRequestStub = jest.spyOn(request, 'get').mockClear().mockImplementation();
      getRequestStub.onCall(0)
        .callsFake(async (opts) => {
          if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
            return {
              "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
              "displayName": "sitecollection-name",
              "name": "sitecollection-name",
              "createdDateTime": "2021-03-09T20:56:00Z",
              "lastModifiedDateTime": "2021-03-09T20:56:01Z",
              "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
            };
          }
          throw 'Invalid request';
        });

      getRequestStub.onCall(1)
        .callsFake(async (opts) => {
          if ((opts.url as string).indexOf('/permissions') > -1) {
            return {
              "value": [
                {
                  "id": "aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                  "grantedToIdentities": [
                    {
                      "application": {
                        "displayName": "Foo",
                        "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
                      }
                    }
                  ]
                },
                {
                  "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
                  "grantedToIdentities": [
                    {
                      "application": {
                        "displayName": "TeamsBotDemo5",
                        "id": "cca00169-d38b-462f-a3b4-f3566b162f2d"
                      }
                    }
                  ]
                }
              ]
            };
          }

          throw 'Invalid request';
        });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/permissions') > -1) {
          return {
            "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
            "roles": [
              "write"
            ],
            "grantedToIdentities": [
              {
                "application": {
                  "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
                }
              }
            ]
          };
        }
        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
          permission: "write",
          appDisplayName: "Foo",
          output: "json"
        }
      });
      assert(loggerLogSpy.calledWith({
        "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
        "roles": [
          "write"
        ],
        "grantedToIdentities": [
          {
            "application": {
              "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
            }
          }
        ]
      }));
    }
  );

  it('Updates an application permission to the site by id', async () => {
    const getRequestStub = jest.spyOn(request, 'get').mockClear().mockImplementation();
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        id: "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
        output: "json"
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
          }
        }
      ]
    }));
  });
});
