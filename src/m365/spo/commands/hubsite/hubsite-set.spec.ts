import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './hubsite-set.js';

describe(commands.HUBSITE_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HUBSITE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates the title of the specified hub site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">Sales</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "Description", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "Description",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "https://contoso.com/logo.png",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "Sales"
    }));
  });

  it('updates the description of the specified hub site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">All things sales</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { description: 'All things sales', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "All things sales",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "https://contoso.com/logo.png",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "Sales"
    }));
  });

  it('updates the logo URL of the specified hub site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">https://contoso.com/logo.png</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { logoUrl: 'https://contoso.com/logo.png', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "All things sales",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "https://contoso.com/logo.png",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "Sales"
    }));
  });

  it('updates the title, description and logo URL of the specified hub site (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">Sales</Parameter></SetProperty><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">https://contoso.com/logo.png</Parameter></SetProperty><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">All things sales</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, title: 'Sales', description: 'All things sales', logoUrl: 'https://contoso.com/logo.png', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "All things sales",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "https://contoso.com/logo.png",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "Sales"
    }));
  });

  it('escapes XML in user input', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">&lt;Sales&gt;</Parameter></SetProperty><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">&lt;https://contoso.com/logo.png&gt;</Parameter></SetProperty><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">&lt;All things sales&gt;</Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "<All things sales>", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "<https:\u002f\u002fcontoso.com\u002flogo.png>", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "<Sales>"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, title: '<Sales>', description: '<All things sales>', logoUrl: '<https://contoso.com/logo.png>', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "<All things sales>",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "<https://contoso.com/logo.png>",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "<Sales>"
    }));
  });

  it('allows resetting hub site title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String"></Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "Description", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": ""
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "Description",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "https://contoso.com/logo.png",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: ""
    }));
  });

  it('allows resetting hub site description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String"></Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "https:\u002f\u002fcontoso.com\u002flogo.png", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { description: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "https://contoso.com/logo.png",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "Sales"
    }));
  });

  it('allows resetting hub site logo URL', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query><SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String"></Parameter></SetProperty><Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">255a50b2-527f-4413-8485-57f4c17a24d1</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": null, "TraceCorrelationId": "3623429e-9057-5000-fcf8-b970da061512"
          }, 34, {
            "IsNull": false
          }, 36, {
            "IsNull": false
          }, 37, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.HubSiteProperties", "Description": "All things sales", "ID": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "LogoUrl": "", "SiteId": "\/Guid(255a50b2-527f-4413-8485-57f4c17a24d1)\/", "SiteUrl": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fSales", "Title": "Sales"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { logoUrl: '', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } });
    assert(loggerLogSpy.calledWith({
      Description: "All things sales",
      ID: "255a50b2-527f-4413-8485-57f4c17a24d1",
      LogoUrl: "",
      SiteId: "255a50b2-527f-4413-8485-57f4c17a24d1",
      SiteUrl: "https://contoso.sharepoint.com/sites/Sales",
      Title: "Sales"
    }));
  });

  it('correctly handles API error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
              "ErrorMessage": "Invalid URL: Logo.", "ErrorValue": null, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPFieldValidationException"
            }, "TraceCorrelationId": "7420429e-a097-5000-fcf8-bab3f3683799"
          }
        ]);
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { logoUrl: 'Logo', id: '255a50b2-527f-4413-8485-57f4c17a24d1' } } as any),
      new CommandError('Invalid URL: Logo.'));
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc', title: 'Sales' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if no property to update specified', async () => {
    const actual = await command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id and title specified', async () => {
    const actual = await command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', title: 'Sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if id and description specified', async () => {
    const actual = await command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', description: 'All things sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if id and logoUrl specified', async () => {
    const actual = await command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', logoUrl: 'https://contoso.com/logo.png' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options specified', async () => {
    const actual = await command.validate({ options: { id: '255a50b2-527f-4413-8485-57f4c17a24d1', title: 'Sales', description: 'All things sales', logoUrl: 'https://contoso.com/logo.png' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
