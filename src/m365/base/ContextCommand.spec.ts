import assert from 'assert';
import fs from 'fs';
import { CommandError } from '../../Command.js';
import { telemetry } from '../../telemetry.js';
import { jestUtil } from '../../utils/jestUtil.js';
import { Hash } from '../../utils/types.js';
import ContextCommand from './ContextCommand.js';

class MockCommand extends ContextCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public mockSaveContextInfo(contextInfo: Hash) {
    this.saveContextInfo(contextInfo);
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('ContextCommand', () => {
  let cmd: MockCommand;
  const contextInfo: Hash = {};

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
  });

  beforeEach(() => {
    cmd = new MockCommand();
  });

  afterEach(() => {
    jestUtil.restore([
      telemetry.trackEvent,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('logs an error if an error occurred while reading the .m365rc.json',
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation().throws(new Error('An error has occurred'));

      assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error reading .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
    }
  );

  it(`logs an error if the .m365rc.json file contents couldn't be parsed`,
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('{');

      let errorMessage;
      try {
        JSON.parse('{');
      }
      catch (err: any) {
        errorMessage = err;
      }

      assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError(`Error reading .m365rc.json: ${errorMessage}. Please add context info to .m365rc.json manually.`));
    }
  );

  it(`logs an error if the content can't be written to the .m365rc.json file`,
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(JSON.stringify({
        "context": {}
      }));
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation().throws(new Error('An error has occurred'));

      assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
    }
  );

  it(`creates the .m365rc.json file if it doesn't exist and saves context info`,
    () => {
      let fileContents: string | undefined;
      let filePath: string | undefined;

      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation((_, contents) => {
        filePath = _.toString();
        fileContents = contents as string;
      });

      cmd.mockSaveContextInfo(contextInfo);

      assert.strictEqual(filePath, '.m365rc.json');
      assert.strictEqual(fileContents, JSON.stringify({
        context: {}
      }, null, 2));
    }
  );

  it(`adds the context info to the existing .m365rc.json file`, () => {
    let fileContents: string | undefined;
    let filePath: string | undefined;

    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(JSON.stringify({}));
    jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation((_, contents) => {
      filePath = _.toString();
      fileContents = contents as string;
    });

    cmd.mockSaveContextInfo(contextInfo);

    assert.strictEqual(filePath, '.m365rc.json');
    assert.strictEqual(fileContents, JSON.stringify({
      context: {}
    }, null, 2));
  });

  it(`doesn't initiate context when it's already present`, () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue(JSON.stringify({
      "context": {}
    }));
    const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear();

    cmd.mockSaveContextInfo(contextInfo);

    assert(fsWriteFileSyncSpy.notCalled);
  });

  it(`doesn't save context info in the .m365rc.json file when there was an error reading file contents`,
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation().throws(new Error());
      const fsWriteFileSyncSpy = jest.spyOn(fs, 'writeFileSync').mockClear();

      assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error reading .m365rc.json: Error. Please add context info to .m365rc.json manually.'));
      assert(fsWriteFileSyncSpy.notCalled);
    }
  );

  it(`doesn't save context info in the .m365rc.json file when there was error writing file contents`,
    () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
      jest.spyOn(fs, 'writeFileSync').mockClear().mockImplementation().throws(new Error('An error has occurred'));

      assert.throws(() => cmd.mockSaveContextInfo(contextInfo), new CommandError('Error writing .m365rc.json: Error: An error has occurred. Please add context info to .m365rc.json manually.'));
    }
  );
});