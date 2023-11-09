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
import { GraphFileDetails } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './file-sharinglink-remove.js';

describe(commands.FILE_SHARINGLINK_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/demo';
  const fileUrl = '/sites/demo/Shared Documents/document.docx';
  const fileId = 'daebb04b-a773-4baa-b1d1-3625418e3234';
  const id = 'U1BEZW1vIFZpc2l0b3Jz';

  const fileInformationResponse: GraphFileDetails = {
    SiteId: '9798e615-a586-455e-8486-84913f492c49',
    VroomDriveID: 'b!FeaYl4alXkWEhoSRP0ksSSOaj9osSfFPqj5bQNdluvlwfL79GNVISZZCf6nfB3vY',
    VroomItemID: '01A5WCPNXHFAS23ZNOF5D3XU2WU7S3I2AU'
  };

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_SHARINGLINK_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, id: id } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the webUrl option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, id: id } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '12345', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('prompts before removing the specified sharing link to a file when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          fileId: fileId,
          id: id
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified sharing link to a file when confirm option not passed and prompt not confirmed',
    async () => {
      const deleteSpy = jest.spyOn(request, 'delete').mockClear();
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          fileUrl: fileUrl,
          id: id
        }
      });

      assert(deleteSpy.notCalled);
    }
  );

  it('removes specified sharing link to a file by fileId when prompt confirmed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
          return fileInformationResponse;
        }

        throw 'Invalid request';
      });

      const requestDeleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileInformationResponse.SiteId}/drives/${fileInformationResponse.VroomDriveID}/items/${fileInformationResponse.VroomItemID}/permissions/${id}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          verbose: true,
          webUrl: webUrl,
          fileId: fileId,
          id: id
        }
      });
      assert(requestDeleteStub.called);
    }
  );

  it('removes specified sharing link to a file by URL', async () => {
    const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        return fileInformationResponse;
      }

      throw 'Invalid request';
    });

    const requestDeleteStub = jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${fileInformationResponse.SiteId}/drives/${fileInformationResponse.VroomDriveID}/items/${fileInformationResponse.VroomItemID}/permissions/${id}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        id: id,
        force: true
      }
    });
    assert(requestDeleteStub.called);
  });

  it('throws error when file not found by id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=SiteId,VroomItemId,VroomDriveId`) {
        throw { error: { 'odata.error': { message: { value: 'File Not Found.' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        id: id,
        force: true,
        verbose: true
      }
    } as any), new CommandError(`File Not Found.`));
  });
});