import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './file-checkout-undo.js';

describe(commands.FILE_CHECKOUT_UNDO, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/projects';
  const fileId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const fileUrl = '/sites/projects/shared documents/test.docx';

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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_CHECKOUT_UNDO);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('undoes checkout for file retrieved by fileId when prompt confirmed',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/getFileById('${fileId}')/undocheckout`) {
          return;
        }

        throw 'Invalid request';
      });
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } });
      assert(postStub.called);
    }
  );

  it('undoes checkout for file retrieved by fileUrl', async () => {
    const serverRelativePath = urlUtil.getServerRelativePath(webUrl, fileUrl);
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, force: true, verbose: true } });
    assert(postStub.called);
  });

  it('undoes checkout for file retrieved by site-relative url', async () => {
    const siteRelativeUrl = '/Shared Documents/Test.docx';
    const serverRelativePath = urlUtil.getServerRelativePath(webUrl, siteRelativeUrl);
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, fileUrl: siteRelativeUrl, force: true, verbose: true } });
    assert(postStub.called);
  });

  it('handles error when file is not checked out', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileById('${fileId}')/undocheckout`) {
        throw {
          error: {
            'odata.error': {
              code: '-2147024738, Microsoft.SharePoint.SPFileCheckOutException',
              message: {
                lang: 'en-US',
                value: 'The file "Shared Documents/4.docx" is not checked out.'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, force: true, verbose: true } }), new CommandError('The file "Shared Documents/4.docx" is not checked out.'));
  });

  it('prompts before undoing checkout when confirmation argument not passed',
    async () => {
      await command.action(logger, { options: { webUrl: webUrl, fileId: fileId } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts undoing checkout when prompt not confirmed', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation().resolves();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: webUrl, id: fileId } });
    assert(postStub.notCalled);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'invalid', fileId: fileId } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and fileId is a valid GUID',
    async () => {
      const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
