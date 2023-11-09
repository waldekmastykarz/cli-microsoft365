import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import userGetCommand from '../../../aad/commands/user/user-get.js';
import commands from '../../commands.js';
import command from './meeting-attendancereport-list.js';

describe(commands.MEETING_ATTENDANCEREPORT_LIST, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const meetingId = 'MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy';
  const meetingAttendanceResponse =
    [
      {
        "id": "ae6ddf54-5d48-4448-a7a9-780eee17fa13",
        "totalParticipantCount": 1,
        "meetingStartDateTime": "2022-11-22T22:46:46.981Z",
        "meetingEndDateTime": "2022-11-22T22:47:07.703Z"
      },
      {
        "id": "3fd019cc-6df5-485f-86a0-96838ab98e66",
        "totalParticipantCount": 1,
        "meetingStartDateTime": "2022-11-22T22:45:10.226Z",
        "meetingEndDateTime": "2022-11-22T22:45:22.347Z"
      },
      {
        "id": "04ddf3a5-0c02-4865-928e-9b65d1b33570",
        "totalParticipantCount": 1,
        "meetingStartDateTime": "2022-11-22T22:43:38.052Z",
        "meetingEndDateTime": "2022-11-22T22:44:12.893Z"
      }
    ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  beforeAll(() => {
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
      Cli.executeCommandWithOutput
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_ATTENDANCEREPORT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'totalParticipantCount']);
  });

  it('fails validation when the userId is not a GUID', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('succeeds validation when the userId and meetingId are valid',
    async () => {
      const actual = await command.validate({ options: { meetingId: meetingId, userId: userId } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('retrieves meeting attendace reports correctly for the current user',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(false);

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meetingId}/attendanceReports`) {
          return { value: meetingAttendanceResponse };
        }
        throw 'Invalid request.';
      });

      await command.action(logger, {
        options:
        {
          meetingId: meetingId
        }
      });

      assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
    }
  );

  it('retrieves meeting attendace reports correctly by userId', async () => {
    jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
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

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userName', async () => {
    jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
      }
      throw 'Invalid request.';
    });

    jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { stdout: JSON.stringify({ id: userId }) };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        meetingId: meetingId,
        userName: userName
      }
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userEmail',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

      jest.spyOn(Cli, 'executeCommandWithOutput').mockClear().mockImplementation(async (command): Promise<any> => {
        if (command === userGetCommand) {
          return { stdout: JSON.stringify({ id: userId }) };
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`) {
          return { value: meetingAttendanceResponse };
        }
        throw 'Invalid request.';
      });

      await command.action(logger, {
        options:
        {
          meetingId: meetingId,
          email: userName,
          verbose: true
        }
      });

      assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
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
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token',
    async () => {
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(false);

      await assert.rejects(command.action(logger, { options: { meetingId: meetingId, userId: userId } } as any),
        new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance reports using delegated permissions`));
    }
  );
});
