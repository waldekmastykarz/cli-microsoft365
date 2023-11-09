import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './retentionlabel-list.js';

describe(commands.RETENTIONLABEL_LIST, () => {

  //#region Mocked responses
  const mockResponseArray = [
    {
      "displayName": "Some label",
      "descriptionForAdmins": "",
      "descriptionForUsers": null,
      "isInUse": true,
      "retentionTrigger": "dateCreated",
      "behaviorDuringRetentionPeriod": "retainAsRecord",
      "actionAfterRetentionPeriod": "delete",
      "createdDateTime": "2022-11-03T10:28:15Z",
      "lastModifiedDateTime": "2022-11-03T10:28:15Z",
      "labelToBeApplied": null,
      "defaultRecordBehavior": "startLocked",
      "id": "dc67203a-6cca-4066-b501-903401308f98",
      "retentionDuration": {
        "days": 365
      },
      "createdBy": {
        "user": {
          "id": "b52ffd35-d6fe-4b70-86d8-91cc01d76333",
          "displayName": null
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "b52ffd35-d6fe-4b70-86d8-91cc01d76333",
          "displayName": null
        }
      },
      "dispositionReviewStages": []
    }
  ];

  const mockResponse = {
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#security/labels/retentionLabels",
    "@odata.count": 2,
    "value": mockResponseArray
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    assert.strictEqual(command.name, commands.RETENTIONLABEL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isInUse']);
  });

  it('retrieves retention labels', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        return mockResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(mockResponseArray));
  });

  it('handles error when retrieving retention labels', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/labels/retentionLabels`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred'));
  });
});