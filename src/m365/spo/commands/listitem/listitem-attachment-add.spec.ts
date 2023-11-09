import assert from 'assert';
import { telemetry } from '../../../../telemetry.js';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import fs from 'fs';
import request from '../../../../request.js';
import command from './listitem-attachment-add.js';

describe(commands.LISTITEM_ATTACHMENT_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listId = '236a0f92482d475bba8fd0e4f78555e4';
  const listTitle = 'Test list';
  const listUrl = 'sites/project-x/lists/testlist';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const listItemId = 1;
  const filePath = 'C:\\Temp\\Test.pdf';
  const fileName = 'CLIRocks.pdf';

  const response = { 'FileName': 'CLIRocks.pdf', 'FileNameAsPath': { 'DecodedUrl': 'Testje.pdf' }, 'ServerRelativePath': { 'DecodedUrl': '/Lists/aaaaaa/Attachments/743/Testje.pdf' }, 'ServerRelativeUrl': '/Lists/aaaaaa/Attachments/743/Testje.pdf' };

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation(((_, defaultValue) => defaultValue));
  });

  afterEach(() => {
    jestUtil.restore([,
      fs.existsSync,
      fs.readFileSync,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LISTITEM_ATTACHMENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      const actual = await command.validate({ options: { webUrl: 'invalid', listTitle: listTitle, listItemId: listItemId, filePath: filePath } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the webUrl option is a valid SharePoint site URL and filePath exists',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath } }, commandInfo);
      assert(actual);
    }
  );

  it('fails validation if the listId option is not a valid GUID', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    const actual = await command.validate({ options: { webUrl: webUrl, listId: 'invalid', listItemId: listItemId, filePath: filePath } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listItemId option is not a valid number',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 'invalid', filePath: filePath } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if the listId option is a valid GUID', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath } }, commandInfo);
    assert(actual);
  });

  it('fails validation if filePath does not exist', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(false);
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('adds attachment to listitem in list retrieved by id while specifying fileName',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('content read');
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (args) => {
        if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileName}')`) {
          return response;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName } });
      assert(loggerLogSpy.calledOnceWith(response));
    }
  );

  it('adds attachment to listitem in list retrieved by url while not specifying fileName',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('content read');
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (args) => {
        if (args.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(${listItemId})/AttachmentFiles/add(FileName='${filePath.replace(/^.*[\\\/]/, '')}')`) {
          return response;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, webUrl: webUrl, listUrl: listUrl, listItemId: listItemId, filePath: filePath } });
      assert(loggerLogSpy.calledOnceWith(response));
    }
  );

  it('adds attachment to listitem in list retrieved by url while specifying fileName without extension',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('content read');
      const fileNameWithoutExtension = fileName.split('.')[0];
      const fileNameWithExtension = `${fileNameWithoutExtension}.${filePath.split('.').pop()}`;
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (args) => {
        if (args.url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileNameWithExtension}')`) {
          return response;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, webUrl: webUrl, listTitle: listTitle, listItemId: listItemId, filePath: filePath, fileName: fileNameWithoutExtension } });
      assert(loggerLogSpy.calledOnceWith(response));
    }
  );

  it('handles error when file with specific name already exists', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('content read');
    const error = {
      error: {
        'odata.error': {
          code: '-2130575257, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'The specified name is already in use.\n\nThe document or folder name was not changed.  To change the name to a different value, close this dialog and edit the properties of the document or folder.'
          }
        }
      }
    };
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileName}')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName } }),
      new CommandError(error.error['odata.error'].message.value.split('\n')[0]));
  });

  it('handles API error', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue('content read');
    const error = {
      error: {
        'odata.error': {
          code: '-2130575257, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'An error has occured.'
          }
        }
      }
    };
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (args) => {
      if (args.url === `${webUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})/AttachmentFiles/add(FileName='${fileName}')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, webUrl: webUrl, listId: listId, listItemId: listItemId, filePath: filePath, fileName: fileName } }),
      new CommandError(error.error['odata.error'].message.value));
  });
});
