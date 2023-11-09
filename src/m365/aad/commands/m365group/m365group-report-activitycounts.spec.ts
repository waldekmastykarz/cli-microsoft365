import assert from 'assert';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './m365group-report-activitycounts.js';

describe(commands.M365GROUP_REPORT_ACTIVITYCOUNTS, () => {
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_REPORT_ACTIVITYCOUNTS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets the number of group activities across group workloads for the given period',
    async () => {
      const requestStub: jest.Mock = jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityCounts(period='D7')`) {
          return `
          Report Refresh Date,Exchange Emails Received,Yammer Messages Posted,Yammer Messages Read,Yammer Messages Liked,Report Date,Report Period
          2019-10-13,,,,,2019-10-13,7
          2019-10-13,,,,,2019-10-12,7
          `;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { period: 'D7' } });
      assert.strictEqual(requestStub.mock.lastCall[0].url, "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityCounts(period='D7')");
      assert.strictEqual(requestStub.mock.lastCall[0].headers["accept"], 'application/json;odata.metadata=none');
    }
  );
});
