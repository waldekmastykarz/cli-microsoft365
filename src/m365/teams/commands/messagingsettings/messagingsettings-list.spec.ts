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
import command from './messagingsettings-list.js';

describe(commands.MESSAGINGSETTINGS_LIST, () => {
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGINGSETTINGS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists messaging settings for a Microsoft Team', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return {
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900" } });
    assert(loggerLogSpy.calledWith({
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }));
  });

  it('lists messaging settings for a Microsoft Team (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return {
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: true } });
    assert(loggerLogSpy.calledWith({
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }));
  });

  it('correctly handles error when listing messaging settings', async () => {
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

    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900" } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when a valid teamId is specified', async () => {
    const actual = await command.validate({
      options: {
        teamId: '2609af39-7775-4f94-a3dc-0dd67657e900'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists all properties for output json', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return {
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", output: 'json' } });
    assert(loggerLogSpy.calledWith({
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }));
  });
});
