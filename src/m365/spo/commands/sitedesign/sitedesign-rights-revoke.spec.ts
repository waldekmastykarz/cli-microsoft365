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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitedesign-rights-revoke.js';

describe(commands.SITEDESIGN_RIGHTS_REVOKE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  beforeAll(() => {
    jest.spyOn(auth, 'restoreAuth').mockClear().mockImplementation().resolves();
    jest.spyOn(telemetry, 'trackEvent').mockClear().mockReturnValue();
    jest.spyOn(pid, 'getProcessName').mockClear().mockReturnValue('');
    jest.spyOn(session, 'getId').mockClear().mockReturnValue('');
    jest.spyOn(spo, 'getRequestDigest').mockClear().mockImplementation().resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    promptOptions = undefined;
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_RIGHTS_REVOKE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('revokes access to the specified site design from the specified principal without prompting for confirmation when confirm option specified',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
          JSON.stringify(opts.data) === JSON.stringify({
            id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            principalNames: ['PattiF']
          })) {
          return {
            "odata.null": true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF' } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('revokes access to the specified site design from the specified principals',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
          JSON.stringify(opts.data) === JSON.stringify({
            id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            principalNames: ['PattiF', 'AdeleV']
          })) {
          return {
            "odata.null": true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF,AdeleV' } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('revokes access to the specified site design from the principals specified via email',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
          JSON.stringify(opts.data) === JSON.stringify({
            id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            principalNames: ['PattiF@contoso.com', 'AdeleV@contoso.com']
          })) {
          return {
            "odata.null": true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF@contoso.com,AdeleV@contoso.com' } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('revokes access to the specified site design from the specified principals separated with an extra space',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
          JSON.stringify(opts.data) === JSON.stringify({
            id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            principalNames: ['PattiF@contoso.com', 'AdeleV@contoso.com']
          })) {
          return {
            "odata.null": true
          };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF@contoso.com, AdeleV@contoso.com' } });
      assert(loggerLogSpy.notCalled);
    }
  );

  it('prompts before revoking access to the specified site design when confirm option not passed',
    async () => {
      await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('aborts revoking access to the site design when prompt not confirmed',
    async () => {
      const postSpy = jest.spyOn(request, 'post').mockClear();
      await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
      assert(postSpy.notCalled);
    }
  );

  it('revokes site design access when prompt confirmed', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation().resolves();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
    assert(postStub.called);
  });

  it('correctly handles error when site script not found', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(new Error('File Not Found.'));

    await assert.rejects(command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF' } } as any), new CommandError('File Not Found.'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteDesignId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying principals', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--principals') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--force') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteDesignId: 'abc', principals: 'PattiF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the all required parameters are specified',
    async () => {
      const actual = await command.validate({ options: { siteDesignId: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a', principals: 'PattiF' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
