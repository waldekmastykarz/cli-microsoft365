import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './user-password-validate.js';

describe(commands.USER_PASSWORD_VALIDATE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

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
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.USER_PASSWORD_VALIDATE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('password is too short', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365"
        })) {
        return {
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_too_short",
              "validationPassed": false,
              "message": "Password is too short, length: 6."
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'cli365' } });
    assert(loggerLogSpy.calledWith({
      "isValid": false,
      "validationResults": [
        {
          "ruleName": "password_too_short",
          "validationPassed": false,
          "message": "Password is too short, length: 6."
        }
      ]
    }));
  });

  it('password complexity is not met', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365password"
        })) {
        return {
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_not_meet_complexity",
              "validationPassed": false,
              "message": "Password does not meet complexity requirements."
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'cli365password' } });
    assert(loggerLogSpy.calledWith({
      "isValid": false,
      "validationResults": [
        {
          "ruleName": "password_not_meet_complexity",
          "validationPassed": false,
          "message": "Password does not meet complexity requirements."
        }
      ]
    }));
  });

  it('password is too weak', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "MyP@ssW0rd"
        })) {
        return {
          "isValid": false,
          "validationResults": [
            {
              "ruleName": "password_banned",
              "validationPassed": false,
              "message": "Password is too weak and can not be used."
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'MyP@ssW0rd' } });
    assert(loggerLogSpy.calledWith({
      "isValid": false,
      "validationResults": [
        {
          "ruleName": "password_banned",
          "validationPassed": false,
          "message": "Password is too weak and can not be used."
        }
      ]
    }));
  });

  it('password meets all requirements', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/beta/users/validatePassword' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "password": "cli365P@ssW0rd"
        })) {
        return {
          "isValid": true,
          "validationResults": [
            {
              "ruleName": "AllChecks",
              "validationPassed": true,
              "message": "Password meets all validation requirements."
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, { options: { password: 'cli365P@ssW0rd' } });
    assert(loggerLogSpy.calledWith({
      "isValid": true,
      "validationResults": [
        {
          "ruleName": "AllChecks",
          "validationPassed": true,
          "message": "Password meets all validation requirements."
        }
      ]
    }));
  });

  it('correctly handles error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      "error": {
        "code": "Error",
        "message": "An error has occurred",
        "innerError": {
          "request-id": "9b0df954-93b5-4de9-8b99-43c204a9acf8",
          "date": "2021-12-08T18:56:48"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { password: 'cli365P@ssW0rd079654' } } as any),
      new CommandError(`An error has occurred`));
  });
});
