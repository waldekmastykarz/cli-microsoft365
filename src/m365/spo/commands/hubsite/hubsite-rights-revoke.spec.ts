import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './hubsite-rights-revoke.js';

describe(commands.HUBSITE_RIGHTS_REVOKE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let promptOptions: any;

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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_RIGHTS_REVOKE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('revokes rights to join the specified hub site without prompting for confirmation when confirm option specified',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
            }, 13, {
              "IsNull": false
            }
          ]);
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin', force: true } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('revokes rights to join the specified hub site without prompting for confirmation when confirm option specified (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
            }, 13, {
              "IsNull": false
            }
          ]);
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin', force: true } });
      assert(loggerLogToStderrSpy.called);
    }
  );

  it('prompts before revoking the rights when confirm option not passed',
    async () => {
      await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts revoking rights when prompt not confirmed', async () => {
    const postSpy = jest.spyOn(request, 'post').mockClear();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });
    await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin' } });
    assert(postSpy.notCalled);
  });

  it('revokes rights when prompt confirmed', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(() => Promise.resolve(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
      }, 13, {
        "IsNull": false
      }
    ])));
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });
    await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin' } });
    assert(postStub.called);
  });

  it('escapes XML in user input', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales&gt;</Parameter><Parameter Type="Array"><Object Type="String">admin&gt;</Object></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
          }, 13, {
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales>', principals: 'admin>', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes rights from the specified principals', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object><Object Type="String">user</Object></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
          }, 13, {
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin,user', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes rights from the specified principals (email)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="Array"><Object Type="String">admin@contoso.onmicrosoft.com</Object><Object Type="String">user@contoso.onmicrosoft.com</Object></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
          }, 13, {
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin@contoso.onmicrosoft.com,user@contoso.onmicrosoft.com', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes rights from the specified principals separated with an extra space',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/Sales</Parameter><Parameter Type="Array"><Object Type="String">admin</Object><Object Type="String">user</Object></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7310.1205", "ErrorInfo": null, "TraceCorrelationId": "71b9439e-800c-5000-b613-208e0afff564"
            }, 13, {
              "IsNull": false
            }
          ]);
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin, user', force: true } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('correctly handles API error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPFieldValidationException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin', force: true } } as any),
      new CommandError('An error has occurred.'));
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { hubSiteUrl: 'abc', principals: 'admin' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all parameters are valid', async () => {
    const actual = await command.validate({ options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all parameters are valid (multiple principals)',
    async () => {
      const actual = await command.validate({ options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin,user' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when all parameters are valid (multiple principals separated with an extra space)',
    async () => {
      const actual = await command.validate({ options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin, user' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when all parameters are valid (multiple principals with email address)',
    async () => {
      const actual = await command.validate({ options: { hubSiteUrl: 'https://contoso.sharepoint.com/sites/Sales', principals: 'admin@contoso.onmicrosoft.com,user@contoso.onmicrosoft.com' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
