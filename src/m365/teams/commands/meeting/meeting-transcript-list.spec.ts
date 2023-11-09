import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './meeting-transcript-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.MEETING_TRANSCRIPT_LIST, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const email = 'user@tenant.com';
  const meetingId = 'MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy';
  const meetingTranscriptsResponse =
    [
      {
        "id": "MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh",
        "createdDateTime": "2021-09-17T06:09:24.8968037Z"
      },
      {
        "id": "MSMjMCMjMzAxNjNhYTctNWRmZi00MjM3LTg5MGQtNWJhYWZjZTZhNWYw",
        "createdDateTime": "2021-09-16T18:58:58.6760692Z"
      },
      {
        "id": "MSMjMCMjNzU3ODc2ZDYtOTcwMi00MDhkLWFkNDItOTE2ZDNmZjkwZGY4",
        "createdDateTime": "2021-09-16T18:56:00.9038309Z"
      }
    ];

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
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
      accessToken.isAppOnlyAccessToken,
      request.get,
      Cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_TRANSCRIPT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'createdDateTime']);
  });

  it('fails validation when the userId is not a GUID', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userName is not valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userName: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('succeeds validation when the userId and meetingId are valid',
    async () => {
      const actual = await command.validate({ options: { meetingId: meetingId, userId: userId } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('succeeds validation when the userName and meetingId are valid',
    async () => {
      const actual = await command.validate({ options: { meetingId: meetingId, userName: userName } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('succeeds validation when the email and meetingId are valid',
    async () => {
      const actual = await command.validate({ options: { meetingId: meetingId, email: email } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation when the userId and email and userName are given',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: { meetingId: meetingId, userId: userId, userName: userName, email: email } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation when given email is not valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, email: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId and email are given', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId, email: email } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId and userName are given', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('retrieves transcript list correctly for the given meetingId for the current user',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(false);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts`) {
          return { value: meetingTranscriptsResponse };
        }
        throw 'Invalid request.';
      });

      await command.action(logger, {
        options:
        {
          meetingId: meetingId
        }
      });

      assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
    }
  );

  it('retrieves meeting transcript list correctly by meetingId and userID',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/onlineMeetings/${meetingId}/transcripts`) {
          return { value: meetingTranscriptsResponse };
        }
        throw 'Invalid request.';
      });

      await command.action(logger, {
        options:
        {
          meetingId: meetingId,
          userId: userId
        }
      });

      assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
    }
  );

  it('retrieves meeting transcript list correctly by meetingId and userName',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/users/${userName}/onlineMeetings/${meetingId}/transcripts`) {
          return { value: meetingTranscriptsResponse };
        }
        throw 'Invalid request.';
      });

      await command.action(logger, {
        options:
        {
          meetingId: meetingId,
          userName: userName
        }
      });

      assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
    }
  );

  it('retrieves meeting transcript list correctly by meetingId and email',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(email)}'&$select=id`) {
          return {
            value: [
              {
                id: userId
              }]
          };
        }

        if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/onlineMeetings/${meetingId}/transcripts`) {
          return { value: meetingTranscriptsResponse };
        }

        throw 'Invalid request';
      });

      await command.action(logger, {
        options:
        {
          meetingId: meetingId,
          email: email,
          verbose: true
        }
      });

      assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
    }
  );

  it('correctly handles error when throwing request', async () => {
    const errorMessage = 'An error has occurred';

    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: { verbose: true, meetingId: meetingId } } as any),
      new CommandError(errorMessage));
  });

  it('correctly handles error when options are missing', async () => {
    jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting transcripts list using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(false);

      await assert.rejects(command.action(logger, { options: { meetingId: meetingId, userId: userId } } as any),
        new CommandError(`The options 'userId', 'userName' and 'email' cannot be used while retrieving meeting transcripts using delegated permissions`));
    }
  );
});