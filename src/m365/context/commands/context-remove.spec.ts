import assert from 'assert';
import fs from 'fs';
import { Cli } from '../../../cli/Cli.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import { telemetry } from '../../../telemetry.js';
import { jestUtil } from '../../../utils/jestUtil.js';
import commands from '../commands.js';
import command from './context-remove.js';

describe(commands.REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let promptOptions: any;

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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      fs.unlinkSync,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the context from the .m365rc.json file when confirm option not passed',
    async () => {
      await command.action(logger, {
        options: {
          debug: false
        }
      });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it(`removes the .m365rc.json file if it exists and only consists of the context parameter`,
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
        context: {}
      }));
      const unlinkSyncStub = jest.spyOn(fs, 'unlinkSync').mockClear().mockImplementation(_ => { });
      await command.action(logger, { options: { debug: true, force: true } });

      assert(unlinkSyncStub.called);
    }
  );

  it(`removes the context info from the existing .m365rc.json file`,
    async () => {
      let fileContents: string | undefined;
      let filePath: string | undefined;

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
        { continue: true }
      ));

      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
        "apps": [
          {
            "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
            "name": "CLI app"
          }
        ],
        "context": {}
      }));
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation((_, contents) => {
        filePath = _.toString();
        fileContents = contents as string;
      });

      await command.action(logger, { options: { debug: true } });

      assert.strictEqual(filePath, '.m365rc.json');
      assert.strictEqual(fileContents, JSON.stringify({
        apps: [
          {
            "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
            "name": "CLI app"
          }
        ]
      }, null, 2));
    }
  );

  it(`handles an error when reading file contents fails`, async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true, force: true } }), new CommandError(`Error reading .m365rc.json: Error: An error has occurred. Please remove context info from .m365rc.json manually.`));
  });

  it(`handles an error when writing file contents fails`, async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
      "apps": [
        {
          "appId": "e23d235c-fcdf-45d1-ac5f-24ab2ee0695d",
          "name": "CLI app"
        }
      ],
      "context": {}
    }));
    jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation(_ => { throw new Error('An error has occurred'); });

    await assert.rejects(command.action(logger, { options: { debug: true, force: true } }), new CommandError(`Error writing .m365rc.json: Error: An error has occurred. Please remove context info from .m365rc.json manually.`));
  });

  it(`handles an error when removing the file fails`, async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
      "context": {}
    }));
    jest.spyOn(fs, 'unlinkSync').mockClear().mockImplementation(_ => { throw new Error('An error has occurred'); });
    await assert.rejects(command.action(logger, { options: { debug: true, force: true } }), new CommandError(`Error removing .m365rc.json: Error: An error has occurred. Please remove .m365rc.json manually.`));
  });

  it(`doesn't update the context file, if it doesn't contain context information`,
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(_ => JSON.stringify({
        apps: [{
          appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
          name: 'My AAD app'
        }]
      }));
      const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear();

      await command.action(logger, { options: { debug: true, force: true } });
      assert(fsWriteFileSyncSpy.notCalled);
    }
  );
});