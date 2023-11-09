import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
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
import command from './knowledgehub-remove.js';

describe(commands.KNOWLEDGEHUB_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];
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
    auth.service.tenantId = 'abc';
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.data) {
          if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`) {
            return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]);
          }
        }
      }

      throw 'Invalid request';
    });
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
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore(Cli.prompt);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.KNOWLEDGEHUB_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes Knowledge Hub settings from tenant without prompting with confirmation argument',
    async () => {
      await command.action(logger, { options: { force: true } });
      let deleteRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`) {
          deleteRequestIssued = true;
        }
      });

      assert(deleteRequestIssued);
    }
  );

  it('removes Knowledge Hub settings from tenant without prompting with confirmation argument (debug)',
    async () => {
      await command.action(logger, { options: { debug: true, force: true } });
      let deleteRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
          r.headers['X-RequestDigest'] &&
          r.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`) {
          deleteRequestIssued = true;
        }
      });

      assert(deleteRequestIssued);
    }
  );

  it('removes Knowledge Hub settings from tenant when confirmation argument not passed',
    async () => {
      await command.action(logger, { options: { debug: true } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing Knowledge Hub settings from tenant when prompt not confirmed',
    async () => {
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

      await command.action(logger, { options: { debug: true } });
      assert(requests.length === 0);
    }
  );

  it('removes removing Knowledge Hub settings from tenant when prompt confirmed',
    async () => {
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { debug: true } });
    }
  );

  it('correctly handles an error when removing Knowledge Hub settings from tenant',
    async () => {
      jestUtil.restore(request.post);
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return { FormDigestValue: 'abc' };
          }
        }

        if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
          if (opts.headers &&
            opts.headers['X-RequestDigest'] &&
            opts.data) {
            if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`) {
              return JSON.stringify([
                {
                  "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
                    "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.InvalidClientQueryException"
                  }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
                }
              ]);
            }
          }
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { debug: true, force: true } } as any), new CommandError('An error has occurred'));
    }
  );
});
