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
import command from './people-profilecardproperty-remove.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_REMOVE, () => {
  let log: string[];
  let logger: Logger;
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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => {
      return { continue: true };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PEOPLE_PROFILECARDPROPERTY_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not a valid value.', async () => {
    const actual = await command.validate({ options: { name: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the name is set to userPrincipalName.',
    async () => {
      const actual = await command.validate({ options: { name: 'userPrincipalName' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('correctly removes profile card property for userPrincipalName',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/UserPrincipalName`) {
          return;
        }

        throw `Invalid request ${opts.url}`;
      });

      await assert.doesNotReject(command.action(logger, { options: { name: 'userPrincipalName' } }));
    }
  );

  it('correctly removes profile card property for userPrincipalName (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/UserPrincipalName`) {
          return;
        }

        throw `Invalid request ${opts.url}`;
      });

      await assert.doesNotReject(command.action(logger, { options: { name: 'userPrincipalName', debug: true } }));
    }
  );

  it('correctly removes profile card property for fax', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/Fax`) {
        return;
      }

      throw `Invalid request ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, { options: { name: 'fax' } }));
  });

  it('correctly removes profile card property for state with force',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/StateOrProvince`) {
          return;
        }

        throw `Invalid request ${opts.url}`;
      });

      await assert.doesNotReject(command.action(logger, { options: { name: 'StateOrProvince', force: true } }));
    }
  );

  it('fails when the removal runs into a property that is not found',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/UserPrincipalName`) {
          throw {
            "error": {
              "code": "404",
              "message": "Not Found",
              "innerError": {
                "peopleAdminErrorCode": "PeopleAdminItemNotFound",
                "peopleAdminRequestId": "2497e6f6-cd91-8bd8-5c53-361d355a5c41",
                "peopleAdminClientRequestId": "1e7328a0-8c5f-476b-9ae1-c1952e2d3276",
                "date": "2023-11-02T19:31:25",
                "request-id": "1e7328a0-8c5f-476b-9ae1-c1952e2d3276",
                "client-request-id": "1e7328a0-8c5f-476b-9ae1-c1952e2d3276"
              }
            }
          };
        }

        throw `Invalid request ${opts.url}`;
      });

      await assert.rejects(command.action(logger, {
        options: {
          name: 'userPrincipalName'
        }
      }), new CommandError(`Not Found`));
    }
  );
});