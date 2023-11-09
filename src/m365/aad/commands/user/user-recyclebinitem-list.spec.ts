import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './user-recyclebinitem-list.js';

describe(commands.USER_RECYCLEBINITEM_LIST, () => {
  const deletedUsersResponse = [{ "businessPhones": [], "displayName": "John Doe", "givenName": "John Doe", "jobTitle": "Developer", "mail": "john@contoso.com", "mobilePhone": "0476345130", "officeLocation": "Washington", "preferredLanguage": "nl-BE", "surname": "John", "userPrincipalName": "7e06b56615f340138bf879874d52e68ajohn@contoso.com", "id": "7e06b566-15f3-4013-8bf8-79874d52e68a" }];
  const graphGetUrl = 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
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
      odata.getAllItems
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_RECYCLEBINITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'userPrincipalName']);
  });

  it('retrieves deleted users', async () => {
    jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation(async (url) => {
      if (url === graphGetUrl) {
        return deletedUsersResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(deletedUsersResponse));
  });

  it('correctly handles API error', async () => {
    jest.spyOn(odata, 'getAllItems').mockClear().mockImplementation().rejects({
      "error": {
        "code": "Invalid_Request",
        "message": "An error has occurred while processing this request.",
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
          "date": "2018-04-24T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { force: true } } as any),
      new CommandError('An error has occurred while processing this request.'));
  });
});