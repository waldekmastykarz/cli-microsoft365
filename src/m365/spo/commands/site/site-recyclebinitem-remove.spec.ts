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
import command from './site-recyclebinitem-remove.js';

describe(commands.SITE_RECYCLEBINITEM_REMOVE, () => {

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
    assert.strictEqual(command.name, commands.SITE_RECYCLEBINITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { siteUrl: 'foo', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1', force: true } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if ids is not a valid guid array string', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,foo', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all options are passed', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1', force: true } }, commandInfo);
    assert(actual);
  });

  it('prompts before removing the items from the recycle bin when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          siteUrl: 'https://contoso.sharepoint.com',
          ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1'
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
      const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation().resolves();
      await command.action(logger, {
        options: {
          siteUrl: 'https://contoso.sharepoint.com',
          ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1'
        }
      });

      assert(postStub.notCalled);
    }
  );

  it('removes items from the recycle bin with ids and confirm option',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteByIds`) {
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
          ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1',
          force: true
        }
      });
    }
  );

  it('removes items from the recycle bin with ids', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteByIds`) {
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
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1'
      }
    });
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          message: {
            value: 'Specified argument was out of the range of valid values'
          }
        }
      }
    };

    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site/RecycleBin/DeleteByIds`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '85528dee-00d5-4c38-a6ba-e2abace32f63,aecb840f-20e9-4ff8-accf-5df8eaad31a1',
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});