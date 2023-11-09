import assert from 'assert';
import { createRequire } from 'module';
import os from 'os';
import auth from '../../../Auth.js';
import { Cli } from '../../../cli/Cli.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { jestUtil } from '../../../utils/jestUtil.js';
import commands from '../commands.js';
import command from './cli-doctor.js';

const require = createRequire(import.meta.url);
const packageJSON = require('../../../../package.json');

describe(commands.DOCTOR, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(() => Promise.resolve());
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
    auth.service.connected = true;
    jest.spyOn(Cli.getInstance().config, 'all').mockClear().mockImplementation().value({});
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
      os.platform,
      os.version,
      os.release,
      process.env
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.DOCTOR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves scopes in the diagnostic information about the current environment',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        scp: 'AllSites.FullControl AppCatalog.ReadWrite.All'
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: [],
        scopes: {
          'https://graph.microsoft.com': [
            'AllSites.FullControl',
            'AppCatalog.ReadWrite.All'
          ]
        }
      }));
    }
  );

  it('retrieves scopes from multiple access tokens in the diagnostic information about the current environment',
    async () => {
      const jwt1 = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        scp: 'AllSites.FullControl AppCatalog.ReadWrite.All'
      });
      let jwt64 = Buffer.from(jwt1).toString('base64');
      const accessToken1 = `abc.${jwt64}.def`;

      const jwt2 = JSON.stringify({
        aud: 'https://mydev.sharepoint.com',
        scp: 'TermStore.Read.All'
      });
      jwt64 = Buffer.from(jwt2).toString('base64');
      const accessToken2 = `abc.${jwt64}.def`;

      const jwt3 = JSON.stringify({
        aud: 'https://mydev-admin.sharepoint.com',
        scp: 'TermStore.Read.All'
      });
      jwt64 = Buffer.from(jwt3).toString('base64');
      const accessToken3 = `abc.${jwt64}.def`;

      const jwt4 = JSON.stringify({
        aud: 'https://mydev-my.sharepoint.com',
        scp: 'TermStore.Read.All'
      });
      jwt64 = Buffer.from(jwt4).toString('base64');
      const accessToken4 = `abc.${jwt64}.def`;

      const jwt5 = JSON.stringify({
        aud: 'https://contoso-admin.sharepoint.com',
        scp: 'TermStore.Read.All'
      });
      jwt64 = Buffer.from(jwt5).toString('base64');
      const accessToken5 = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken1}` },
        'https://mydev.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken2}` },
        'https://mydev-admin.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken3}` },
        'https://mydev-my.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken4}` },
        'https://contoso-admin.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken5}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: [],
        scopes: {
          'https://graph.microsoft.com': [
            'AllSites.FullControl',
            'AppCatalog.ReadWrite.All'
          ],
          'https://mydev.sharepoint.com': [
            'TermStore.Read.All'
          ],
          'https://contoso.sharepoint.com': [
            'TermStore.Read.All'
          ]
        }
      }));
    }
  );

  it('retrieves roles in the diagnostic information about the current environment',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All']
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scopes: {}
      }));
    }
  );

  it('retrieves roles from multiple access tokens in the diagnostic information about the current environment',
    async () => {
      const jwt1 = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All']
      });
      let jwt64 = Buffer.from(jwt1).toString('base64');
      const accessToken1 = `abc.${jwt64}.def`;

      const jwt2 = JSON.stringify({
        aud: 'https://mydev.sharepoint.com',
        roles: ['TermStore.Read.All']
      });
      jwt64 = Buffer.from(jwt2).toString('base64');
      const accessToken2 = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken1}` },
        'https://mydev.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken2}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All', 'TermStore.Read.All'],
        scopes: {}
      }));
    }
  );

  it('retrieves roles and scopes in the diagnostic information about the current environment',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scp: 'Sites.Read.All Files.ReadWrite.All'
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scopes: {
          'https://graph.microsoft.com':
            [
              'Sites.Read.All',
              'Files.ReadWrite.All'
            ]
        }
      }));
    }
  );

  it('retrieves diagnostic information about the current environment when there are no roles or scopes available',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: [],
        scp: ''
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });

      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: [],
        scopes: {}
      }));
    }
  );

  it('retrieves diagnostic information about the current environment with auth type Certificate',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All']
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(2);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'Certificate',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scopes: {}
      }));
    }
  );

  it('retrieves tenant information as single when TenantID is a GUID',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All']
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('923d42f0-6d23-41eb-b68d-c036d242654f');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(2);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: { debug: true } });
      assert(loggerLogSpy.calledWith({
        authMode: 'Certificate',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'single',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scopes: {}
      }));
    }
  );

  it('retrieves diagnostic information about the current environment (debug)',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All']
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(2);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: { debug: true } });
      assert(loggerLogSpy.calledWith({
        authMode: 'Certificate',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scopes: {}
      }));
    }
  );

  it('retrieves diagnostic information of the current environment when executing in docker',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        roles: ['Sites.Read.All', 'Files.ReadWrite.All']
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(2);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': 'docker' });

      await command.action(logger, { options: { debug: true } });
      assert(loggerLogSpy.calledWith({
        authMode: 'Certificate',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: 'docker',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
        scopes: {}
      }));
    }
  );

  it('returns empty roles and scopes in diagnostic information when access token is empty',
    async () => {
      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': '' }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(2);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: { debug: true } });
      assert(loggerLogSpy.calledWith({
        authMode: 'Certificate',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: [],
        scopes: {}
      }));
    }
  );


  it('returns empty roles and scopes in diagnostic information when access token is invalid',
    async () => {
      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': 'a.b.c.d' }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(2);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });

      await command.action(logger, { options: { debug: true } });
      assert(loggerLogSpy.calledWith({
        authMode: 'Certificate',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {},
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: [],
        scopes: {}
      }));
    }
  );

  it('retrieves CLI Configuration in the diagnostic information about the current environment',
    async () => {
      const jwt = JSON.stringify({
        aud: 'https://graph.microsoft.com',
        scp: 'AllSites.FullControl AppCatalog.ReadWrite.All'
      });
      const jwt64 = Buffer.from(jwt).toString('base64');
      const accessToken = `abc.${jwt64}.def`;

      jest.spyOn(auth.service, 'accessTokens').mockClear().mockImplementation().value({
        'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
      });
      jest.spyOn(os, 'platform').mockClear().mockReturnValue('win32');
      jest.spyOn(os, 'version').mockClear().mockReturnValue('Windows 10 Pro');
      jest.spyOn(os, 'release').mockClear().mockReturnValue('10.0.19043');
      jest.spyOn(packageJSON, 'version').mockClear().mockImplementation().value('3.11.0');
      jest.spyOn(process, 'version').mockClear().mockImplementation().value('v14.17.0');
      jest.spyOn(auth.service, 'appId').mockClear().mockImplementation().value('31359c7f-bd7e-475c-86db-fdb8c937548e');
      jest.spyOn(auth.service, 'tenant').mockClear().mockImplementation().value('common');
      jest.spyOn(auth.service, 'authType').mockClear().mockImplementation().value(0);
      jest.spyOn(process, 'env').mockClear().mockImplementation().value({ 'CLIMICROSOFT365_ENV': '' });
      jestUtil.restore(Cli.getInstance().config.all);
      jest.spyOn(Cli.getInstance().config, 'all').mockClear().mockImplementation().value({ "showHelpOnFailure": false });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith({
        authMode: 'DeviceCode',
        cliAadAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        cliAadAppTenant: 'common',
        cliEnvironment: '',
        cliVersion: '3.11.0',
        cliConfig: {
          "showHelpOnFailure": false
        },
        nodeVersion: 'v14.17.0',
        os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
        roles: [],
        scopes: {
          'https://graph.microsoft.com': [
            'AllSites.FullControl',
            'AppCatalog.ReadWrite.All'
          ]
        }
      }));
    }
  );
});
