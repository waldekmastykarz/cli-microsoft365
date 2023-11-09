import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './gateway-get.js';

describe(commands.GATEWAY_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

  const gateway: any = {
    id: "1f69e798-5852-4fdd-ab01-33bb14b6e934",
    name: "My_Sample_Gateway",
    type: "Resource",
    publicKey: {
      exponent: "AQAB",
      modulus: "o6j2....cLk="
    }
  };

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
    loggerLogSpy = jest.spyOn(logger, "log").mockClear();
  });

  afterEach(() => {
    jestUtil.restore([request.get]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.strictEqual(command.name, commands.GATEWAY_GET);
  });

  it("has a description", () => {
    assert.notStrictEqual(command.description, null);
  });

  it("fails validation if the id is not valid", async () => {
    const actual = await command.validate(
      {
        options: {
          id: "3eb1a01b"
        }
      },
      commandInfo
    );

    assert.notStrictEqual(actual, true);
  });

  it("passes validation if the id is valid", async () => {
    const actual = await command.validate(
      {
        options: {
          id: "1f69e798-5852-4fdd-ab01-33bb14b6e934"
        }
      },
      commandInfo
    );

    assert.strictEqual(actual, true);
  });

  it("correctly handles error", async () => {
    jest.spyOn(request, "get").mockClear().mockImplementation(() => {
      throw "An error has occurred";
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          id: 'testid'
        }
      }),
      new CommandError("An error has occurred")
    );
  });

  it("should get gateway information for the gateway by id", async () => {
    jest.spyOn(request, "get").mockClear().mockImplementation((opts) => {
      if (
        opts.url ===
        "https://api.powerbi.com" +
        `/v1.0/myorg/gateways/${formatting.encodeQueryParameter(gateway.id)}`
      ) {
        return gateway;
      }
      throw "Invalid request";
    });

    await command.action(logger, {
      options: {
        id: "1f69e798-5852-4fdd-ab01-33bb14b6e934"
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.mock.lastCall;

    assert.strictEqual(call.mock.calls[0].id, "1f69e798-5852-4fdd-ab01-33bb14b6e934");
    assert.strictEqual(call.mock.calls[0].name, "My_Sample_Gateway");
    assert(loggerLogSpy.calledWith(gateway));
  });
});
