import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitedesign-task-remove.js';

describe(commands.SITEDESIGN_TASK_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(spo, 'ensureFormDigest').mockClear().mockImplementation().resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
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
    assert.strictEqual(command.name, commands.SITEDESIGN_TASK_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified site design task without prompting for confirmation when confirm option specified',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RemoveSiteDesignTask`) > -1 &&
          JSON.stringify(opts.data) === JSON.stringify({
            taskId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
          })) {
          return {
            "odata.null": true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { force: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('prompts before removing the specified site design task when confirm option not passed',
    async () => {
      await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing site design task when prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();

      await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
      assert(postSpy.notCalled);
    }
  );

  it('removes the site design task when prompt confirmed', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation().resolves();

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(postStub.called);
  });

  it('supports specifying taskId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { taskId: 'b2307a39-e878-458b-bc90-03bc578531d6', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--force') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the taskId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the taskId is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
