import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './userprofile-set.js';

describe(commands.USERPROFILE_SET, () => {
  let log: any[];
  let logger: Logger;
  const spoUrl = 'https://contoso.sharepoint.com';

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
    auth.service.spoUrl = spoUrl;
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USERPROFILE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates single valued profile property', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`) > -1) {
        return {
          "odata.null": true
        };
      }
      throw 'Invalid request';
    });

    const data: any = {
      'accountName': `i:0#.f|membership|john.doe@mytenant.onmicrosoft.com`,
      'propertyName': 'SPS-JobTitle',
      'propertyValue': 'Senior Developer'
    };

    await command.action(logger, {
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer',
        debug: true
      }
    });
    const lastCall = postStub.mock.lastCall[0];
    assert.strictEqual(JSON.stringify(lastCall.data), JSON.stringify(data));
  });

  it('updates multi valued profile property', async () => {
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty`) > -1) {
        return {
          "odata.null": true
        };
      }
      throw 'Invalid request';
    });

    const data: any = {
      'accountName': `i:0#.f|membership|john.doe@mytenant.onmicrosoft.com`,
      'propertyName': 'SPS-Skills',
      'propertyValues': ['CSS', 'HTML']
    };

    await command.action(logger, {
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-Skills',
        propertyValue: 'CSS, HTML'
      }
    });
    const lastCall = postStub.mock.lastCall[0];
    assert.strictEqual(JSON.stringify(lastCall.data), JSON.stringify(data));
  });

  it('correctly handles error while updating profile property', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
