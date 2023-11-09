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
import command from './mail-send.js';

describe(commands.MAIL_SEND, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

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
    requests = [];
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
    assert.strictEqual(command.name, commands.MAIL_SEND);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Send an email to one recipient (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and from someone (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'someone@contoso.com', verbose: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and from someone', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'someone@contoso.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and from some peoples (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);
        if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'user@contoso.com,someone@consotos.com', verbose: true } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.data) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
  );

  it('Send an email to one recipient and from some peoples', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'user@contoso.com,someone@consotos.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and CC someone (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', cc: 'someone@contoso.com', verbose: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and CC someone', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', cc: 'someone@contoso.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and BCC someone (debug)', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', bcc: 'someone@contoso.com', verbose: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient and BCC someone', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', bcc: 'someone@contoso.com' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('Send an email to one recipient with additional header (debug)',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        requests.push(opts);
        if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
          return '';
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', additionalHeaders: '{"X-Custom": "My Custom Header Value"}', verbose: true } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.data) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
  );

  it('Send an email to one recipient with additional header', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return '';
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', additionalHeaders: '{"X-Custom": "My Custom Header Value"}' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('correctly handles random API error', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(() => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', additionalHeaders: '{"X-Custom": "My Custom Header Value"}' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if at least the webUrl \'to\', \'subject\' and \'body\' are sprecified',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );
});
