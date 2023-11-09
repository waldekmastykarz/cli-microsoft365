import assert from 'assert';
import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import commands from '../../commands.js';
import command from './completion-sh-update.js';

describe(commands.COMPLETION_SH_UPDATE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let generateShCompletionStub: jest.Mock;

  beforeAll(() => {
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockImplementation(() => { });
    jest.spyOn(pid, 'getProcessName').mockClear().mockImplementation(() => '');
    jest.spyOn(session, 'getId').mockClear().mockImplementation(() => '');
    generateShCompletionStub = jest.spyOn(autocomplete, 'generateShCompletion').mockClear().mockImplementation(() => { });
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
  });

  afterEach(() => {
    generateShCompletionStub.mockReset();
  });

  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_SH_UPDATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('builds command completion', async () => {
    await command.action(logger, { options: {} });
    assert(generateShCompletionStub.called);
  });

  it('build command completion (debug)', async () => {
    await command.action(logger, { options: { debug: true } });
    assert(loggerLogToStderrSpy.called);
  });
});
