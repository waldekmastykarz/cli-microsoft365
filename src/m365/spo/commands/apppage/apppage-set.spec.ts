import assert from 'assert';
import auth from "../../../../Auth.js";
import { Cli } from "../../../../cli/Cli.js";
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from "../../../../cli/Logger.js";
import { CommandError } from "../../../../Command.js";
import request from "../../../../request.js";
import { pid } from "../../../../utils/pid.js";
import { session } from "../../../../utils/session.js";
import { jestUtil } from "../../../../utils/jestUtil.js";
import { telemetry } from "../../../../telemetry.js";
import commands from "../../commands.js";
import command from './apppage-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APPPAGE_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  beforeAll(() => {
    cli = Cli.getInstance();
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it("has correct name", () => {
    assert.strictEqual(command.name, commands.APPPAGE_SET);
  });

  it("has a description", () => {
    assert.notStrictEqual(command.description, null);
  });

  it("fails to update the single-part app page if request is rejected",
    async () => {
      jest.spyOn(request, "post").mockClear().mockImplementation(async opts => {
        if (
          (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1 &&
          opts.data.serverRelativeUrl.indexOf("failme")
        ) {
          throw "Failed to update the single-part app page";
        }
        throw 'Invalid request';
      });
      await assert.rejects(command.action(logger,
        {
          options: {
            name: "failme",
            webUrl: "https://contoso.sharepoint.com/",
            webPartData: JSON.stringify({})
          }
        }), new CommandError(`Failed to update the single-part app page`));
    }
  );

  it("Update the single-part app pag", async () => {
    jest.spyOn(request, "post").mockClear().mockImplementation(async opts => {
      if (
        (opts.url as string).indexOf(`_api/sitepages/Pages/UpdateFullPageApp`) > -1
      ) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger,
      {
        options: {
          pageName: "demo",
          webUrl: "https://contoso.sharepoint.com/teams/sales",
          webPartData: JSON.stringify({})
        }
      });
  });

  it("supports specifying name", () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--name") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying webUrl", () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--webUrl") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("supports specifying webPartData", () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf("--webPartData") > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it("fails validation if name not specified", async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webPartData not specified", async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        name: "Contoso.aspx",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webUrl not specified", async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webPartData: JSON.stringify({ abc: "def" }),
        name: "page.aspx"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it("fails validation if webPartData is not a valid JSON string",
    async () => {
      const actual = await command.validate({
        options: {
          name: "Contoso.aspx",
          webUrl: "https://contoso",
          webPartData: "abc"
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it("validation passes on all required options", async () => {
    const actual = await command.validate({
      options: {
        name: "Contoso.aspx",
        webPartData: "{}",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
