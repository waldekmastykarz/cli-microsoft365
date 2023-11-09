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
import command from './retentionevent-get.js';

describe(commands.RETENTIONEVENT_GET, () => {
  const retentionEventId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';
  const retentionEventGetResponse = {
    "displayName": "Employee Termination",
    "description": "This event occurs when an employee is terminated.",
    "eventTriggerDateTime": "2023-02-01T09:16:37Z",
    "lastStatusUpdateDateTime": "2023-02-01T09:21:15Z",
    "createdDateTime": "2023-02-01T09:17:40Z",
    "lastModifiedDateTime": "2023-02-01T09:17:40Z",
    "id": retentionEventId,
    "eventQueries": [
      {
        "queryType": "files",
        "query": "1234"
      },
      {
        "queryType": "messages",
        "query": "Terminate"
      }
    ],
    "eventStatus": {
      "error": null,
      "status": "success"
    },
    "eventPropagationResults": [
      {
        "serviceName": "SharePoint",
        "location": null,
        "status": "none",
        "statusInformation": null
      }
    ],
    "createdBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    },
    "lastModifiedBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    }
  };

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
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct id is entered', async () => {
    const actual = await command.validate({ options: { id: retentionEventId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves retention event by specified id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents/${retentionEventId}`) {
        return retentionEventGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: retentionEventId, verbose: true } });
    assert(loggerLogSpy.calledWith(retentionEventGetResponse));
  });

  it('handles error when retention event by specified id is not found',
    async () => {
      const errorMessage = `Error: The operation couldn't be performed because object '${retentionEventId}' couldn't be found on 'FfoConfigurationSession'.`;
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents/${retentionEventId}`) {
          throw errorMessage;
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, {
        options: {
          id: retentionEventId
        }
      }), new CommandError(errorMessage));
    }
  );
});