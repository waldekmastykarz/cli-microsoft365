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
import command from './file-move.js';

describe(commands.FILE_MOVE, () => {
  const fileName = 'Report.docx';
  const rootUrl = 'https://contoso.sharepoint.com';
  const webUrl = rootUrl + '/sites/project-x';
  const sourceUrl = '/sites/project-x/documents/' + fileName;
  const targetUrl = '/sites/project-y/My Documents';
  const absoluteSourceUrl = rootUrl + sourceUrl;
  const absoluteTargetUrl = rootUrl + targetUrl;
  const sourceId = 'b8cc341b-9c11-4f2d-aa2b-0ce9c18bcba2';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let postStub: jest.Mock;

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

    postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async opts => {
      if (opts.url === `${webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request: ' + opts.url;
    });
  });

  afterEach(() => {
    jestUtil.restore([
      request.post,
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_MOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['targetUrl', 'sourceUrl']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: 'foo', sourceUrl: sourceUrl, targetUrl: targetUrl } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if sourceId is not a valid guid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceId: 'invalid', targetUrl: targetUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if nameConflictBehavior is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceUrl: sourceUrl, targetUrl: targetUrl, nameConflictBehavior: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the sourceId is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, sourceId: sourceId, targetUrl: targetUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL',
    async () => {
      const actual = await command.validate({ options: { webUrl: webUrl, sourceUrl: sourceUrl, targetUrl: targetUrl } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('moves a file correctly when specifying sourceId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `${webUrl}/_api/Web/GetFileById('${sourceId}')?$select=ServerRelativePath`) {
        return {
          ServerRelativePath: {
            DecodedUrl: sourceUrl
          }
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        sourceId: sourceId,
        targetUrl: targetUrl,
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.mock.lastCall[0].data,
      {
        srcPath: {
          DecodedUrl: absoluteSourceUrl
        },
        destPath: {
          DecodedUrl: absoluteTargetUrl + `/${fileName}`
        },
        overwrite: false,
        options: {
          KeepBoth: false,
          ShouldBypassSharedLocks: false,
          RetainEditorAndModifiedOnMove: false
        }
      }
    );
  });

  it('moves a file correctly when specifying sourceUrl with server relative paths',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          sourceUrl: sourceUrl,
          targetUrl: targetUrl
        }
      });

      assert.deepStrictEqual(postStub.mock.lastCall[0].data,
        {
          srcPath: {
            DecodedUrl: absoluteSourceUrl
          },
          destPath: {
            DecodedUrl: absoluteTargetUrl + `/${fileName}`
          },
          overwrite: false,
          options: {
            KeepBoth: false,
            ShouldBypassSharedLocks: false,
            RetainEditorAndModifiedOnMove: false
          }
        }
      );
    }
  );

  it('moves a file correctly when specifying sourceUrl with site relative paths',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          sourceUrl: `/Shared Documents/${fileName}`,
          targetUrl: targetUrl,
          nameConflictBehavior: 'fail'
        }
      });

      assert.deepStrictEqual(postStub.mock.lastCall[0].data,
        {
          srcPath: {
            DecodedUrl: webUrl + `/Shared Documents/${fileName}`
          },
          destPath: {
            DecodedUrl: absoluteTargetUrl + `/${fileName}`
          },
          overwrite: false,
          options: {
            KeepBoth: false,
            ShouldBypassSharedLocks: false,
            RetainEditorAndModifiedOnMove: false
          }
        }
      );
    }
  );

  it('moves a file correctly when specifying sourceUrl with absolute paths',
    async () => {
      await command.action(logger, {
        options: {
          webUrl: webUrl,
          sourceUrl: rootUrl + sourceUrl,
          targetUrl: rootUrl + targetUrl,
          nameConflictBehavior: 'replace'
        }
      });

      assert.deepStrictEqual(postStub.mock.lastCall[0].data,
        {
          srcPath: {
            DecodedUrl: absoluteSourceUrl
          },
          destPath: {
            DecodedUrl: absoluteTargetUrl + `/${fileName}`
          },
          overwrite: true,
          options: {
            KeepBoth: false,
            ShouldBypassSharedLocks: false,
            RetainEditorAndModifiedOnMove: false
          }
        }
      );
    }
  );

  it('moves a file correctly when specifying various options', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        sourceUrl: sourceUrl,
        targetUrl: targetUrl,
        newName: 'Report-old.docx',
        nameConflictBehavior: 'rename',
        retainEditorAndModified: true,
        bypassSharedLock: true
      }
    });

    assert.deepStrictEqual(postStub.mock.lastCall[0].data,
      {
        srcPath: {
          DecodedUrl: absoluteSourceUrl
        },
        destPath: {
          DecodedUrl: absoluteTargetUrl + '/Report-old.docx'
        },
        overwrite: false,
        options: {
          KeepBoth: true,
          ShouldBypassSharedLocks: true,
          RetainEditorAndModifiedOnMove: true
        }
      }
    );
  });

  it('handles error correctly when moving a file', async () => {
    const error = {
      error: {
        'odata.error': {
          message: {
            lang: 'en-US',
            value: 'File Not Found.'
          }
        }
      }
    };

    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: webUrl,
        sourceId: sourceId,
        targetUrl: targetUrl
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
