import assert from 'assert';
import auth from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { jestUtil } from '../../../utils/jestUtil.js';
import commands from '../commands.js';
import command from './flow-disable.js';

describe(commands.DISABLE, () => {
  let log: string[];
  let logger: Logger;

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
  });

  afterEach(() => {
    jestUtil.restore([
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DISABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('disables the specified flow (debug)', async () => {
    const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } });
    assert.strictEqual(postStub.mock.lastCall[0].url, 'https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/stop?api-version=2016-11-01');
  });

  it('disables the specified flow as admin', async () => {
    const postStub: jest.Mock = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments`) > -1) {

        return;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } }));
    assert.strictEqual(postStub.mock.lastCall[0].url, 'https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/stop?api-version=2016-11-01');
  });

  it('correctly handles no environment found', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      "error": {
        "code": "EnvironmentAccessDenied",
        "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } } as any),
      new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`));
  });

  it('correctly handles Flow not found', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      "error": {
        "code": "ConnectionAuthorizationFailed",
        "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12' } } as any),
      new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'.`));
  });

  it('correctly handles Flow not found (as admin)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      "error": {
        "code": "FlowNotFound",
        "message": "Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'."
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12', asAdmin: true } } as any),
      new CommandError(`Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'.`));
  });

  it('correctly handles API OData error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } } as any),
      new CommandError('An error has occurred'));
  });
});
