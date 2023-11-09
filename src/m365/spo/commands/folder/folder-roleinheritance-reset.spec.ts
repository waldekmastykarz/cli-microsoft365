import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './folder-roleinheritance-reset.js';

describe(commands.FOLDER_ROLEINHERITANCE_RESET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderUrl = 'Shared Documents/TestFolder';
  const rootFolderUrl = '/Shared Documents';

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
    assert.strictEqual(command.name, commands.FOLDER_ROLEINHERITANCE_RESET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: folderUrl, force: true } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if webUrl and folderUrl are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, folderUrl: folderUrl, force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before resetting role inheritance for the folder when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          folderUrl: folderUrl
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts resetting role inheritance for the folder when confirm option is not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          folderUrl: folderUrl
        }
      });

      assert(postSpy.notCalled);
    }
  );

  it('resets role inheritance on folder by site-relative URL (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20Documents%2FTestFolder')/ListItemAllFields/resetroleinheritance`) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          verbose: true,
          webUrl: webUrl,
          folderUrl: folderUrl,
          force: true
        }
      });
    }
  );

  it('resets role inheritance on folder by site-relative URL when prompt confirmed',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='%2Fsites%2Fproject-x%2FShared%20Documents%2FTestFolder')/ListItemAllFields/resetroleinheritance`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          folderUrl: folderUrl
        }
      });
    }
  );

  it('resets role inheritance on root folder URL of a library when prompt confirmed',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `${webUrl}/_api/web/GetList('%2Fsites%2Fproject-x%2FShared%20Documents')/resetroleinheritance`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          webUrl: webUrl,
          folderUrl: rootFolderUrl
        }
      });
    }
  );

  it('correctly handles error when resetting folder role inheritance',
    async () => {
      const errorMessage = 'Cannot find resource';
      jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
        error: {
          'odata.error': {
            message: {
              value: errorMessage
            }
          }
        }
      });

      await assert.rejects(command.action(logger, {
        options: {
          debug: true,
          webUrl: webUrl,
          folderUrl: folderUrl,
          force: true
        }
      }), new CommandError(errorMessage));
    }
  );
});