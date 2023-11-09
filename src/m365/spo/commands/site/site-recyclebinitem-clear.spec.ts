import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './site-recyclebinitem-clear.js';

describe(commands.SITE_RECYCLEBINITEM_CLEAR, () => {

  let log: any[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

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
      request.post,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_RECYCLEBINITEM_CLEAR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { siteUrl: 'foo', force: true } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the webUrl option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', force: true } }, commandInfo);
      assert(actual);
    }
  );

  it('prompts before removing the items from the recycle bin when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          siteUrl: 'https://contoso.sharepoint.com'
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the items from the recycle bin when confirm option not passed and prompt not confirmed',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/RecycleBin/DeleteAll`) {
          return {
            'odata.null': true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          siteUrl: 'https://contoso.sharepoint.com'
        }
      });

      assert(postStub.notCalled);
    }
  );

  it('removes all items from the first-stage recycle bin with confirm option',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/RecycleBin/DeleteAll`) {
          return {
            'odata.null': true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          verbose: true,
          siteUrl: 'https://contoso.sharepoint.com',
          force: true
        }
      });
    }
  );

  it('removes all items from the first-stage recycle bin without confirmation',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/web/RecycleBin/DeleteAll`) {
          return {
            'odata.null': true
          };
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));

      await command.action(logger, {
        options: {
          siteUrl: 'https://contoso.sharepoint.com'
        }
      });

      assert(postStub.called);
    }
  );

  it('removes all items from the second-stage recycle bin with confirm option',
    async () => {
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteAllSecondStageItems`) {
          return {
            'odata.null': true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          verbose: true,
          siteUrl: 'https://contoso.sharepoint.com',
          secondary: true,
          force: true
        }
      });

      assert(postStub.called);
    }
  );

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          message: {
            value: "The files cannot be removed from the second-stage recycle bin."
          }
        }
      }
    };

    jest.spyOn(request, 'post').mockClear().mockImplementation(async () => {
      return error;
    });

    await assert.rejects(
      command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com', force: true } } as any),
      new CommandError(error.error['odata.error'].message.value)
    );
  });
});
