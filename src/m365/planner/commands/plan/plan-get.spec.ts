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
import command from './plan-get.js';

describe(commands.PLAN_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  const validId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTitle = 'Plan name';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const validRosterId = 'FeMZFDoK8k2oWmuGE-XFHZcAEwtn';
  const invalidOwnerGroupId = 'Invalid GUID';

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const planResponse = {
    "id": validId,
    "title": validTitle
  };

  const planDetailsResponse = {
    "sharedWith": {},
    "categoryDescriptions": {}
  };

  const outputResponse = {
    ...planResponse,
    ...planDetailsResponse
  };

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
    assert.strictEqual(command.name, commands.PLAN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner', '@odata.etag']);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id specified', async () => {
    const actual = await command.validate({
      options: {
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when title and valid ownerGroupId specified',
    async () => {
      const actual = await command.validate({
        options: {
          title: validTitle,
          ownerGroupId: validOwnerGroupId
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when title and valid ownerGroupName specified',
    async () => {
      const actual = await command.validate({
        options: {
          title: validTitle,
          ownerGroupName: validOwnerGroupName
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when rosterId specified', async () => {
    const actual = await command.validate({
      options: {
        rosterId: validRosterId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly get planner plan with given id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return planDetailsResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        id: validId
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with given title and ownerGroupId',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
          return { "value": [planResponse] };
        }

        if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
          return planDetailsResponse;
        }

        throw `Invalid request ${opts.url}`;
      });

      const options: any = {
        title: validTitle,
        ownerGroupId: validOwnerGroupId
      };

      await command.action(logger, { options: options });
      assert(loggerLogSpy.calledWith(outputResponse));
    }
  );

  it('correctly get planner plan with given title and ownerGroupName',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
          return singleGroupResponse;
        }

        if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
          return { "value": [planResponse] };
        }

        if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
          return planDetailsResponse;
        }

        throw `Invalid request ${opts.url}`;
      });

      const options: any = {
        title: validTitle,
        ownerGroupName: validOwnerGroupName
      };

      await command.action(logger, { options: options } as any);
      assert(loggerLogSpy.calledWith(outputResponse));
    }
  );

  it('correctly get planner plan with given rosterId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans`) {
        return { "value": [planResponse] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return planDetailsResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    const options: any = {
      rosterId: validRosterId
    };

    await command.action(logger, { options: options });
    assert(loggerLogSpy.calledWith(outputResponse));
  });


  it('correctly handles no plan found with given ownerGroupId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return { "value": [] };
      }

      throw `Invalid request ${opts.url}`;
    });

    const options: any = {
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    await assert.rejects(command.action(logger, { options: options } as any));
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error(`Planner plan with id '${validId}' was not found.`));

    await assert.rejects(command.action(logger, { options: { id: validId } }), new CommandError(`Planner plan with id '${validId}' was not found.`));
  });
});
