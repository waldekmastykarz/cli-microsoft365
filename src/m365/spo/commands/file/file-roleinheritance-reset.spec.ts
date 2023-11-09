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
import spoFileGetCommand from './file-get.js';
import command from './file-roleinheritance-reset.js';

describe(commands.FILE_ROLEINHERITANCE_RESET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const fileUrl = '/sites/project-x/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';

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
      Cli.prompt,
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ROLEINHERITANCE_RESET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, force: true } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '0cd891ef-afce-4e55-b836-fce03286cccf', force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before resetting role inheritance for the file when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          fileId: fileId
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts resetting role inheritance for the file when confirm option is not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          fileId: fileId
        }
      });

      assert(postSpy.notCalled);
    }
  );

  it('resets role inheritance on file by relative URL (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/resetroleinheritance`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        force: true
      }
    });
  });

  it('resets role inheritance on file by Id when prompt confirmed',
    async () => {
      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === spoFileGetCommand) {
          return ({
            stdout: '{"LinkingUri": "https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866","Name": "Test1.docx","ServerRelativeUrl": "/sites/project-x/documents/Test1.docx","UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"}'
          });
        }

        throw new CommandError('Unknown case');
      });

      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/resetroleinheritance`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          webUrl: 'https://contoso.sharepoint.com/sites/project-x',
          fileId: fileId
        }
      });
    }
  );

  it('correctly handles error when resetting file role inheritance',
    async () => {
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
      jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(error);

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          webUrl: webUrl,
          fileUrl: fileUrl,
          force: true
        }
      }), new CommandError(error.error['odata.error'].message.value));
    }
  );
});
