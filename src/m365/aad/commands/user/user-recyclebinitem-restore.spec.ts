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
import command from './user-recyclebinitem-restore.js';

describe(commands.USER_RECYCLEBINITEM_RESTORE, () => {
  const validUserId = 'd839826a-81bf-4c38-8f80-f150d11ce6c7';
  const userResponse = {
    id: 'cc9467d2-00f8-4ce7-b0c5-11a401936f08',
    businessPhones: [
      '+1 309 555 0104'
    ],
    displayName: 'John Doe',
    givenName: 'John',
    jobTitle: 'Developer',
    mail: 'john@contoso.com',
    mobilePhone: null,
    officeLocation: '19/2109',
    preferredLanguage: 'John Doe',
    surname: 'Doe',
    userPrincipalName: 'john@contoso.com'
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
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_RECYCLEBINITEM_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('restores the user from the recycle bin', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/${validUserId}/restore`) {
        return userResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validUserId, verbose: true } });
    assert(loggerLogSpy.calledWith(userResponse));
  });

  it('correctly handles API error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      error: {
        error: {
          code: 'Request_ResourceNotFound',
          message: `Resource '${validUserId}' does not exist or one of its queried reference-property objects are not present.`,
          innerError: {
            'request-id': '9b0df954-93b5-4de9-8b99-43c204a8aaf8',
            date: '2018-04-24T18:56:48'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: validUserId } } as any),
      new CommandError(`Resource '${validUserId}' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});