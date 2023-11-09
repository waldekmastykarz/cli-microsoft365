import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './bucket-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.BUCKET_SET, () => {
  let cli: Cli;
  const validBucketId = 'vncYUXCRBke28qMLB-d4xJcACtNz';
  const validBucketName = 'Bucket name';
  const validOrderHint = '8585513699476931356P;';
  const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
  const validPlanTitle = 'Plan title';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const invalidOwnerGroupId = 'Invalid GUID';
  const validRosterId = 'RuY-PSpdw02drevnYDTCJpgAEfoI';

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const multipleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      },
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const singlePlanResponse = {
    "value": [
      {
        "id": validPlanId,
        "title": validPlanTitle
      }
    ]
  };

  const singleBucketByNameResponse = {
    "value": [
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
        "name": validBucketName,
        "id": validBucketId
      }
    ]
  };

  const singleBucketByIdResponse = {
    "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
    "name": validBucketName,
    "id": validBucketId
  };

  const multipleBucketByNameResponse = {
    "value": [
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
        "name": validBucketName,
        "id": validBucketId
      },
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
        "name": validBucketName,
        "id": validBucketId
      }
    ]
  };

  const planResponse = {
    value: [{
      id: validPlanId,
      title: validPlanTitle
    }]
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.patch,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.BUCKET_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation id when id and plan details are specified',
    async () => {
      const actual = await command.validate({
        options: {
          id: validBucketId,
          planId: validPlanId
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation when owner group id is not a guid', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group name',
    async () => {
      const actual = await command.validate({
        options: {
          name: validBucketName,
          planId: validPlanId,
          ownerGroupName: validOwnerGroupName
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation when plan id is used with owner group id', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when roster id is used with owner group name',
    async () => {
      const actual = await command.validate({
        options: {
          name: validBucketName,
          rosterId: validRosterId,
          ownerGroupName: validOwnerGroupName
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation when roster id is used with owner group id',
    async () => {
      const actual = await command.validate({
        options: {
          name: validBucketName,
          rosterId: validRosterId,
          ownerGroupId: validOwnerGroupId
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation when new name or order hint is not specified',
    async () => {
      const actual = await command.validate({
        options: {
          id: validBucketId
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId,
        newName: 'New name'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with name', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        newName: 'New name'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when no groups found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError(`The specified group '${validOwnerGroupName}' does not exist.`));
  });

  it('fails validation when multiple groups found', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return multipleGroupResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError("Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('fails validation when no buckets found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planId: validPlanId
      }
    }), new CommandError(`The specified bucket ${validBucketName} does not exist`));
  });

  it('fails validation when multiple buckets found', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return multipleBucketByNameResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planId: validPlanId
      }
    }), new CommandError("Multiple buckets with name 'Bucket name' found. Found: vncYUXCRBke28qMLB-d4xJcACtNz."));
  });

  it('handles selecting single result when multiple buckets with the specified name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
          return singleGroupResponse;
        }
        if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
          return singlePlanResponse;
        }
        if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
          return multipleBucketByNameResponse;
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
          return;
        }

        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves(singleBucketByNameResponse.value[0]);

      await assert.doesNotReject(command.action(logger, {
        options: {
          name: validBucketName,
          planTitle: validPlanTitle,
          ownerGroupName: validOwnerGroupName,
          newName: 'New bucket name',
          orderHint: validOrderHint
        }
      }));
    }
  );

  it('correctly updates bucket by id', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return singleBucketByIdResponse;
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        id: validBucketId,
        newName: validBucketName
      }
    }));
  });

  it('correctly updates bucket by name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return singleGroupResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return singleBucketByNameResponse;
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        newName: 'New bucket name',
        orderHint: validOrderHint
      }
    }));
  });

  it('correctly updates bucket by name with rosterId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans`) {
        return planResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return singleBucketByNameResponse;
      }

      throw 'Invalid Request';
    });

    const patchStub = jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        name: validBucketName,
        rosterId: validRosterId,
        newName: 'New bucket name',
        orderHint: validOrderHint
      }
    });

    assert(patchStub.called);
  });

  it('correctly updates bucket by name with group ID', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return singleBucketByNameResponse;
      }

      throw 'Invalid request';
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: validOwnerGroupId,
        newName: 'New bucket name',
        orderHint: validOrderHint
      }
    }));
  });
});
