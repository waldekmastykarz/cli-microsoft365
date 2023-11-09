import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './folder-add.js';

describe(commands.FOLDER_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let stubPostResponses: any;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;

    stubPostResponses = (addResp: any = null) => {
      return jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web/folders') > -1) {
          if (addResp) {
            throw addResp;
          }
          else {
            return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "abc", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/abc", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };
          }
        }

        throw 'Invalid request';
      });
    };
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
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should correctly handle folder add reject request', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    stubPostResponses(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('should correctly handle folder add success request', async () => {
    stubPostResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    });
    assert(loggerLogSpy.mock.lastCall.calledWith({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "abc", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/abc", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" }));
  });

  it('should correctly pass params to request', async () => {
    const request: jest.Mock = stubPostResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    });
    assert(request.calledWith({
      url: `https://contoso.sharepoint.com/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter('/Shared Documents/abc')}')`,
      headers:
        { accept: 'application/json;odata=nometadata' },
      responseType: 'json'
    }));
  });

  it('should correctly pass params to request (sites/test1)', async () => {
    const request: jest.Mock = stubPostResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        parentFolderUrl: 'Shared Documents',
        name: 'abc'
      }
    });
    assert(request.calledWith({
      url: `https://contoso.sharepoint.com/sites/test1/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter('/sites/test1/Shared Documents/abc')}')`,
      headers:
        { accept: 'application/json;odata=nometadata' },
      responseType: 'json'
    }));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', parentFolderUrl: '/Shared Documents', name: 'My Folder' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the webUrl option is a valid SharePoint site URL and parentFolderUrl specified',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', parentFolderUrl: '/Shared Documents', name: 'My Folder' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
