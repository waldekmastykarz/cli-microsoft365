import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './tenant-recyclebinitem-list.js';

describe(commands.TENANT_RECYCLEBINITEM_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

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
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_RECYCLEBINITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['DaysRemaining', 'DeletionTime', 'Url']);
  });

  it('handles client.svc promise error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('handles error while getting tenant recycle bin', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.ServerException"
            }, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc"
          }
        ]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
  it('includes all properties for json output', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19527.12077", "ErrorInfo": null, "TraceCorrelationId": "85bb2b9f-5099-2000-af64-2c100126d549"
          }, 54, {
            "IsNull": false
          }, 56, {
            "IsNull": false
          }, 57, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
              }
            ]
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } });
    assert(loggerLogSpy.calledWith([
      {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
      }, {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
      }
    ]));
  });

  it('lists the tenant recyclebin items (debug)', async () => {

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.19527.12077", "ErrorInfo": null, "TraceCorrelationId": "85bb2b9f-5099-2000-af64-2c100126d549"
          }, 54, {
            "IsNull": false
          }, 56, {
            "IsNull": false
          }, 57, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPODeletedSitePropertiesEnumerable", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fClassicThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,4,3,893)\/", "SiteId": "\/Guid(7db536da-792b-4be7-b9b6-194778905606)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fClassicThrowaway", "UserCodeMaximumLevel": 0
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "85bb2b9f-5099-2000-af64-2c100126d549|908bed80-a04a-4433-b4a0-883d9847d110:c7d25483-6785-4e76-8b22-9c57c0b70134\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fModernThrowaway", "DaysRemaining": 92, "DeletionTime": "\/Date(2020,0,15,11,40,58,90)\/", "SiteId": "\/Guid(38fb96c1-8e1d-4d24-ad8d-e57cb9b1749e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fModernThrowaway", "UserCodeMaximumLevel": 300
              }
            ]
          }
        ]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert.strictEqual(loggerLogSpy.mock.lastCall[0][0]["DaysRemaining"], 92);
    assert.deepEqual(loggerLogSpy.mock.lastCall[0][0]["DeletionTime"], new Date(2020, 0, 15, 11, 4, 3, 893));
    assert.strictEqual(loggerLogSpy.mock.lastCall[0][0]["Url"], 'https://contoso.sharepoint.com/sites/ClassicThrowaway');
    assert.strictEqual(loggerLogSpy.mock.lastCall[0][1].DaysRemaining, 92);
    assert.deepEqual(loggerLogSpy.mock.lastCall[0][1].DeletionTime, new Date(2020, 0, 15, 11, 40, 58, 90));
    assert.strictEqual(loggerLogSpy.mock.lastCall[0][1].Url, 'https://contoso.sharepoint.com/sites/ModernThrowaway');
  });

  it('handles tenant recyclebin timeout', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7407.1202", "ErrorInfo": { "ErrorMessage": "Timed out" }, "TraceCorrelationId": "2df74b9e-c022-5000-1529-309f2cd00843"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant"
          }
        ]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Timed out'));
  });
});
