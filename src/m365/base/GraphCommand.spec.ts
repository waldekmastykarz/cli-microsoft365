import assert from 'assert';
import { telemetry } from '../../telemetry.js';
import auth from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { jestUtil } from '../../utils/jestUtil.js';
import GraphCommand from './GraphCommand.js';

class MockCommand extends GraphCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('GraphCommand', () => {
  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
  });

  afterEach(() => {
    jestUtil.restore(auth.restoreAuth);
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('correctly reports an error while restoring auth info', async () => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation(async () => { throw 'An error has occurred'; });
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('doesn\'t execute command when error occurred while restoring auth info',
    async () => {
      jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().rejects(new Error('An error has occurred'));
      const command = new MockCommand();
      const logger: Logger = {
        log: async () => { },
        logRaw: async () => { },
        logToStderr: async () => { }
      };
      const commandCommandActionSpy = jest.spyOn(command, 'commandAction').mockClear();
      await assert.rejects(command.action(logger, { options: {} }));
      assert(commandCommandActionSpy.notCalled);
    }
  );

  it('doesn\'t execute command when not logged in', async () => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    auth.service.connected = false;
    const commandCommandActionSpy = jest.spyOn(command, 'commandAction').mockClear();
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('executes command when logged in', async () => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    auth.service.connected = true;
    const commandCommandActionSpy = jest.spyOn(command, 'commandAction').mockClear();
    await command.action(logger, { options: {} });
    assert(commandCommandActionSpy.called);
  });

  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://graph.microsoft.com');
  });
});
