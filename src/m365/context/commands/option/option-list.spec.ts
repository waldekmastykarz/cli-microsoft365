import assert from 'assert';
import fs from 'fs';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './option-list.js';

describe(commands.OPTION_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
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
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.OPTION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles an error when reading file content fails', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true } }), new CommandError(`Error reading .m365rc.json: Error: An error has occurred. Please retrieve context options from .m365rc.json manually.`));
  });

  it(`retrieves context info options from the existing .m365rc.json file`,
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
        "apps": [
          {
            "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
            "name": "CLI app"
          }
        ],
        "context": {
          "listName": "listNameValue"
        }
      }));

      await command.action(logger, { options: { verbose: true } });
      assert(loggerLogSpy.calledWith({ "listName": "listNameValue" }));
    }
  );

  it('handles an error when context is not present in the .m365rc.json file',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
        "apps": [
          {
            "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
            "name": "CLI app"
          }
        ]
      }));

      await assert.rejects(command.action(logger, { options: { debug: true, name: 'listName', force: true } }), new CommandError(`No context present`));
    }
  );

});