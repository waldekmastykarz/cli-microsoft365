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
import command from './tenant-recyclebinitem-restore.js';

describe(commands.TENANT_RECYCLEBINITEM_RESTORE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
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
  });

  afterEach(() => {
    (command as any).currentContext = undefined;
    jestUtil.restore([
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_RECYCLEBINITEM_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the url option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr' } }, commandInfo);
      assert(actual);
    }
  );

  it(`restores deleted site collection from the tenant recycle bin, without waiting for completion`,
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/SPOInternalUseOnly.Tenant/RestoreDeletedSite`) > -1) {
          if (opts.headers &&
            JSON.stringify(opts.data) === JSON.stringify({
              siteUrl: 'https://contoso.sharepoint.com/sites/hr'
            })) {
            return "{\"HasTimedout\":false,\"IsComplete\":true,\"PollingInterval\":15000}";
          }
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr' } });
    }
  );

  it('handles error when the site to restore is not found', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/SPOInternalUseOnly.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers &&
          JSON.stringify(opts.data) === JSON.stringify({
            siteUrl: 'https://contoso.sharepoint.com/sites/hr'
          })) {
          throw "{\"odata.error\":{\"code\":\"-2147024809, System.ArgumentException\",\"message\":{\"lang\":\"en-US\",\"value\":\"Unable to find the deleted site: https://contoso.sharepoint.com/sites/hr.\"}}}";
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/hr' } } as any), new CommandError("{\"odata.error\":{\"code\":\"-2147024809, System.ArgumentException\",\"message\":{\"lang\":\"en-US\",\"value\":\"Unable to find the deleted site: https://contoso.sharepoint.com/sites/hr.\"}}}"));
  });
});
