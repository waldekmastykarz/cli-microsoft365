import assert from 'assert';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './sensitivitylabel-list.js';

describe(commands.SENSITIVITYLABEL_LIST, () => {
  const userId = '59f80e08-24b1-41f8-8586-16765fd830d3';
  const userName = 'john.doe@contoso.com';
  const sensitivityLabelListResponse = {
    "value": [
      {
        "id": "6f4fb2db-ecf4-4279-94ba-23d059bf157e",
        "name": "Unrestricted",
        "description": "",
        "color": "",
        "sensitivity": 0,
        "tooltip": "Information either intended for general distribution, or which would not have any impact on the organization if it were to be distributed.",
        "isActive": true,
        "isAppliable": true,
        "contentFormats": [
          "file",
          "email"
        ],
        "hasProtection": false,
        "parent": null
      }
    ]
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
    jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(false);
  });

  afterEach(() => {
    jestUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SENSITIVITYLABEL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userId defined', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userName defined', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'isActive']);
  });

  it('retrieves sensitivity labels that the current logged in user has access to',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/me/security/informationProtection/sensitivityLabels`) {
          return sensitivityLabelListResponse;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: {} });
      assert(loggerLogSpy.calledWith(sensitivityLabelListResponse.value));
    }
  );

  it('retrieves sensitivity labels that the specific user has access to by its Id',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/security/informationProtection/sensitivityLabels`) {
          return sensitivityLabelListResponse;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { userId: userId } });
      assert(loggerLogSpy.calledWith(sensitivityLabelListResponse.value));
    }
  );

  it('retrieves sensitivity labels that the specific user has access to by its UPN',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/users/${userName}/security/informationProtection/sensitivityLabels`) {
          return sensitivityLabelListResponse;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { userName: userName } });
      assert(loggerLogSpy.calledWith(sensitivityLabelListResponse.value));
    }
  );

  it('throws an error when using application permissions and no option is specified',
    async () => {
      jestUtil.restore(accessToken.isAppOnlyAccessToken);
      jest.spyOn(accessToken, 'isAppOnlyAccessToken').mockClear().mockReturnValue(true);

      await assert.rejects(command.action(logger, {
        options: {}
      }), new CommandError(`Specify at least 'userId' or 'userName' when using application permissions.`));
    }
  );

  it('handles error when retrieving sensitivity labels', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/security/informationProtection/sensitivityLabels`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred'));
  });
});