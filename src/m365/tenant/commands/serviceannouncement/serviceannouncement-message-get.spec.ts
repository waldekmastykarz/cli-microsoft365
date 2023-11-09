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
import commands from "../../commands.js";
import command from './serviceannouncement-message-get.js';

describe(commands.SERVICEANNOUNCEMENT_MESSAGE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  const testId = 'MC001337';
  const testIncorrectId = '123456';

  const resResourceNotExist = {
    "error": {
      "code": "UnknownError",
      "message": "{\"code\":\"forbidden\",\"message\":\"{\\u0022error\\u0022:\\u0022Resource doesn\\\\u0027t exist for the tenant. ActivityId: b2307a39-e878-458b-bc90-03bc578531d6. Learn more: https://docs.microsoft.com/en-us/graph/api/resources/service-communications-api-overview?view=graph-rest-beta\\\\u0026preserve-view=true.\\u0022}\"}",
      "innerError": {
        "date": "2022-01-22T15:01:15",
        "request-id": "b2307a39-e878-458b-bc90-03bc578531d6",
        "client-request-id": "b2307a39-e878-458b-bc90-03bc578531d6"
      }
    }
  };

  const resMessage = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#admin/serviceAnnouncement/messages/$entity",
    "startDateTime": "2021-02-01T19:23:04Z",
    "endDateTime": "2022-01-31T08:00:00Z",
    "lastModifiedDateTime": "2021-02-01T19:24:37.837Z",
    "title": "Service reminder: Skype for Business Online retires in 6 months",
    "id": "MC001337",
    "category": "planForChange",
    "severity": "normal",
    "tags": [
      "User impact",
      "Admin impact"
    ],
    "isMajorChange": false,
    "actionRequiredByDateTime": "2021-07-31T07:00:00Z",
    "services": [
      "Skype for Business"
    ],
    "expiryDateTime": null,
    "hasAttachments": false,
    "viewPoint": null,
    "details": [
      {
        "name": "BlogLink",
        "value": "https://techcommunity.microsoft.com/t5/microsoft-teams-blog/skype-for-business-online-will-retire-in-12-months-plan-for-a/ba-p/1554531"
      },
      {
        "name": "ExternalLink",
        "value": "https://docs.microsoft.com/microsoftteams/skype-for-business-online-retirement"
      }
    ],
    "body": {
      "contentType": "html",
      "content": "<p>Originally announced in MC219641 (July '20), as Microsoft Teams has become the core communications client for Microsoft 365, this is a reminder the Skype for Business Online service will <a href=\"https://techcommunity.microsoft.com/t5/microsoft-teams-blog/skype-for-business-online-will-retire-in-12-months-plan-for-a/ba-p/1554531\" target=\"_blank\">retire July 31, 2021</a>. At that point, access to the service will end.</p><p>Please click Additional Information to learn more.</p>"
    }
  };

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
    assert.strictEqual(command.name, commands.SERVICEANNOUNCEMENT_MESSAGE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if incorrect message ID is provided', async () => {
    const actual = await command.validate({
      options: {
        id: testIncorrectId
      }
    }, commandInfo);
    assert.strictEqual(actual, `${testIncorrectId} is not a valid message ID`);
  });

  it('passes validation if correct message ID is provided', async () => {
    const actual = await command.validate({
      options: {
        id: testId
      }
    }, commandInfo);
    assert(actual);
  });

  it('correctly retrieves service update message', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testId}`) {
        return resMessage;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: testId
      }
    });
    assert.strictEqual(loggerLogSpy.calledWith(resMessage), true);
    assert.strictEqual(loggerLogSpy.mock.lastCall[0].id, testId);
    assert.strictEqual(loggerLogSpy.callCount, 1);
  });

  it('correctly retrieves service update message (debug)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testId}`) {
        return resMessage;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: testId
      }
    });
    assert.strictEqual(loggerLogSpy.calledWith(resMessage), true);
    assert.strictEqual(loggerLogSpy.mock.lastCall[0].id, testId);
    assert.strictEqual(loggerLogSpy.callCount, 1);
  });

  it('fails when the message does not exist for the tenant', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testIncorrectId}`) {
        throw resResourceNotExist;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: testIncorrectId } } as any), new CommandError(resResourceNotExist.error.message));
  });

  it('lists all properties for output json', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/${testId}`) {
        return resMessage;
      }

      throw 'Invalid request';
    });


    await command.action(logger, {
      options:
      {
        id: testId,
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith(resMessage));
  });
});
