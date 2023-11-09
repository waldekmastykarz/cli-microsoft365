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
import command from './people-profilecardproperty-get.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_GET, () => {
  const profileCardPropertyName = 'customAttribute1';

  //#region Mocked responses
  const response = {
    directoryPropertyName: profileCardPropertyName,
    annotations: [
      {
        displayName: 'Cost center',
        localizations: [
          {
            languageTag: 'nl-NL',
            displayName: 'Kostencentrum'
          }
        ]
      }
    ]
  };
  //#endregion

  let log: any[];
  let loggerLogSpy: jest.SpyInstance;
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
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
    assert.strictEqual(command.name, commands.PEOPLE_PROFILECARDPROPERTY_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when name is invalid', async () => {
    const actual = await command.validate({ options: { name: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name is valid', async () => {
    const actual = await command.validate({ options: { name: profileCardPropertyName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name is valid with different capitalization',
    async () => {
      const actual = await command.validate({ options: { name: 'cUstoMATTriBUtE1' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('gets profile card property information', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('gets profile card property information for text output', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    const textOutput = {
      directoryPropertyName: profileCardPropertyName,
      displayName: response.annotations[0].displayName,
      ['displayName ' + response.annotations[0].localizations[0].languageTag]: response.annotations[0].localizations[0].displayName
    };

    await command.action(logger, { options: { name: profileCardPropertyName, output: 'text' } });
    assert(loggerLogSpy.calledOnceWith(textOutput));
  });

  it('handles error when profile card property does not exist', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({
      response: {
        status: 404
      }
    });

    await assert.rejects(command.action(logger, { options: { name: profileCardPropertyName } } as any),
      new CommandError(`Profile card property '${profileCardPropertyName}' does not exist.`));
  });

  it('handles unexpected API error', async () => {
    const errorMessage = 'Something went wrong';
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { name: profileCardPropertyName } } as any),
      new CommandError(errorMessage));
  });
});