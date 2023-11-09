import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './schemaextension-remove.js';

describe(commands.SCHEMAEXTENSION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let loggerLogToStderrSpy: jest.SpyInstance;
  let promptOptions: any;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    auth.service.connected = true;
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes schema extension', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('removes schema extension (debug)', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'exttyee4dv5_MySchemaExtension', force: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('prompts before removing schema extension when confirmation argument not passed',
    async () => {
      await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts removing schema extension when prompt not confirmed',
    async () => {
      jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects('Invalid request');
      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

      await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
    }
  );

  it('removes schema extension when prompt confirmed', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects({ error: 'An error has occurred' });

    await assert.rejects(command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    jest.spyOn(request, 'delete').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
