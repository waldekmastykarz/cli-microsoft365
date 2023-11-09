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
import commands from '../../commands.js';
import command from './tab-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TAB_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
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
    promptOptions = undefined;
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TAB_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when valid channelId, teamId and id is specified',
    async () => {
      const actual = await command.validate({
        options: {
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation if the teamId , channelId and id are not provided',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {

        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if the channelId is not valid channelId', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: 'invalid',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the id is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });


  it('prompts before removing the specified tab when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('prompts before removing the specified tab when confirm option not passed (debug)',
    async () => {
      await command.action(logger, {
        options: {
          debug: true,
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified tab when confirm option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'delete').mockClear();
      await command.action(logger, {
        options: {
          debug: true,
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
        }
      });
      assert(postSpy.notCalled);
    }
  );

  it('aborts removing the specified tab when confirm option not passed and prompt not confirmed (debug)',
    async () => {
      const postSpy = jest.spyOn(request, 'delete').mockClear();
      await command.action(logger, {
        options: {
          debug: true,
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
        }
      });
      assert(postSpy.notCalled);
    }
  );

  it('removes the specified tab by id when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`tabs/d66b8110-fcad-49e8-8159-0d488ddb7656`) > -1) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          debug: true,
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
        }
      });
    }
  );


  it('removes the specified tab without prompting when confirmed specified (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`tabs/d66b8110-fcad-49e8-8159-0d488ddb7656`) > -1) {
          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options: {
          debug: true,
          channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
          teamId: '00000000-0000-0000-0000-000000000000',
          id: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
          force: true
        }
      });
    }
  );

  it('handles error correctly', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
