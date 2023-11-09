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
import command from './schemaextension-set.js';

describe(commands.SCHEMAEXTENSION_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: jest.SpyInstance;
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
    loggerLogToStderrSpy = jest.spyOn(logger, 'logToStderr').mockClear();
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.patch
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates schema extension', async () => {
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return {
          "id": "ext6kguklm2_TestSchemaExtension",
          "description": "Test Description",
          "targetTypes": [
            "Group"
          ],
          "status": "InDevelopment",
          "owner": "b07a45b3-f7b7-489b-9269-da6f3f93dff0",
          "properties": [
            {
              "name": "MyInt",
              "type": "Integer"
            },
            {
              "name": "MyString",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        status: 'Available',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    });
    assert.strictEqual(log.length, 0);
  });

  it('updates schema extension (debug)', async () => {
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        status: 'Available',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    });
    assert(loggerLogToStderrSpy.calledWith("Schema extension successfully updated."));
  });

  it('updates schema extension (verbose)', async () => {
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/schemaExtensions/ext6kguklm2_TestSchemaExtension`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        id: 'ext6kguklm2_TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        status: 'Available'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('handles error correctly', async () => {
    jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if the owner is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'invalid',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if no update information is specified', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if properties is not valid JSON string', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: 'foobar'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if properties have no valid type', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Foo"},{"name":"MyString","type":"String"}]'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if a specified property has missing type', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt"},{"name":"MyString","type":"String"}]'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if a specified property has missing name', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"type":"Integer"},{"name":"MyString","type":"String"}]'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if properties JSON string is not an array',
    async () => {
      const actual = await command.validate({
        options: {
          id: 'TestSchemaExtension',
          description: 'Test Description',
          owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
          targetTypes: 'Group',
          properties: '{}'
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if status is not valid', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: 'Test Description',
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        status: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required parameters are set and at least one property to update (description) is specified',
    async () => {
      const actual = await command.validate({
        options: {
          id: 'TestSchemaExtension',
          owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
          description: 'test'
        }
      }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if the property type is Binary', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Binary"}]'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is Boolean', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Boolean"}]'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is DateTime', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"DateTime"}]'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is Integer', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"Integer"}]'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the property type is String', async () => {
    const actual = await command.validate({
      options: {
        id: 'TestSchemaExtension',
        description: null,
        owner: 'b07a45b3-f7b7-489b-9269-da6f3f93dff0',
        targetTypes: 'Group',
        properties: '[{"name":"MyInt","type":"String"}]'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
