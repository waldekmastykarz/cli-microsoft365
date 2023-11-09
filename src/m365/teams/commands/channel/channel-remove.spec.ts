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
import command from './channel-remove.js';

describe(commands.CHANNEL_REMOVE, () => {
  const id = '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype';
  const name = 'channelName';
  const teamId = 'd66b8110-fcad-49e8-8159-0d488ddb7656';
  const teamName = 'Team Name';

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when valid id & teamId is specified', async () => {
    const actual = await command.validate({
      options: {
        id: id,
        teamId: teamId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name & teamName is specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Channel Name',
        teamName: teamName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id is not valid', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        id: id
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails to remove channel when channel does not exists', async () => {
    const errorMessage = 'The specified channel does not exist in this Microsoft Teams team';
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamId: teamId,
        name: name,
        force: true
      }
    }), new CommandError(errorMessage));
  });

  it('prompts before removing the specified channel when confirm option not passed (debug)',
    async () => {
      await command.action(logger, {
        options: {
          debug: true,
          id: id,
          teamId: teamId
        }
      });

      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing the specified channel when confirm option not passed and prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'delete').mockClear();

      await command.action(logger, {
        options: {
          debug: true,
          id: id,
          teamId: teamId
        }
      });

      assert(postSpy.notCalled);
    }
  );

  it('fails when team name does not exist', async () => {
    const errorMessage = 'The specified team does not exist';
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": []
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        teamName: teamName,
        force: true
      }
    }), new CommandError(errorMessage));
  });

  it('removes specified channel when id is passed with confirm option',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockReturnValue(Promise.resolve());

      await command.action(logger, {
        options: {
          id: id,
          teamId: teamId,
          force: true
        }
      });
    }
  );

  it('removes the specified channel by id when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(id)}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));

      await command.action(logger, {
        options: {
          debug: true,
          id: id,
          teamId: teamId
        }
      });
    }
  );

  it('removes the specified channel by name and teamName when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'`) {
          return {
            value: [
              {
                "id": teamId,
                "displayName": teamName,
                "resourceProvisioningOptions": ["Team"]
              }
            ]
          };
        }
        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(id)}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          debug: true,
          id: id,
          teamName: teamName
        }
      });
    }
  );


  it('removes the specified channel by name and teamId when prompt confirmed (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'`) {
          return {
            value: [
              {
                "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
                "displayName": "name",
                "description": null,
                "email": "",
                "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6%40thread.skype/%F0%9F%92%A1+Ideas?groupId=d66b8110-fcad-49e8-8159-0d488ddb7656&tenantId=eff8592e-e14a-4ae8-8771-d96d5c549e1c"
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(id)}`) {
          return;
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, {
        options: {
          debug: true,
          name: name,
          teamId: teamId
        }
      });
    }
  );

  it('removes the specified channel by name without prompt', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'`) {
        return {
          value: [
            {
              "id": teamId,
              "displayName": teamName,
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'`) {
        return {
          value: [
            {
              "id": id,
              "displayName": name,
              "description": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6%40thread.skype/%F0%9F%92%A1+Ideas?groupId=d66b8110-fcad-49e8-8159-0d488ddb7656&tenantId=eff8592e-e14a-4ae8-8771-d96d5c549e1c"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(id)}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: name,
        teamName: teamName,
        force: true
      }
    });
  });

  it('correctly handles Microsoft graph error response', async () => {
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
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        id: id,
        teamId: teamId,
        force: true
      }
    }), new CommandError(error.error.message));
  });
});
