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
import command from './room-list.js';

describe(commands.ROOM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;

  const jsonOutput = {
    "value": [
      {
        "id": "3162F1E1-C4C0-604B-51D8-91DA78989EB1",
        "emailAddress": "cf100@contoso.com",
        "displayName": "Conf Room 100",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geoCoordinates": {
          "latitude": 47.6405,
          "longitude": -122.1293
        },
        "phone": "000-000-0000",
        "nickname": "Conf Room",
        "label": "100",
        "capacity": 50,
        "building": "1",
        "floorNumber": 1,
        "isManaged": true,
        "isWheelChairAccessible": false,
        "bookingType": "standard",
        "tags": [
          "bean bags"
        ],
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDevice": "surface hub"
      },
      {
        "id": "3162F1E1-C4C0-604B-51D8-91DA78970B97",
        "emailAddress": "cf200@contoso.com",
        "displayName": "Conf Room 200",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geoCoordinates": {
          "latitude": 47.6405,
          "longitude": -122.1293
        },
        "phone": "000-000-0000",
        "nickname": "Conf Room",
        "label": "200",
        "capacity": 40,
        "building": "2",
        "floorNumber": 2,
        "isManaged": true,
        "isWheelChairAccessible": false,
        "bookingType": "standard",
        "tags": [
          "benches",
          "nice view"
        ],
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDevice": "surface hub"
      }
    ]
  };
  const jsonOutputFilter = {
    "value": [
      {
        "id": "3162F1E1-C4C0-604B-51D8-91DA78970B97",
        "emailAddress": "cf200@contoso.com",
        "displayName": "Conf Room 200",
        "address": {
          "street": "4567 Main Street",
          "city": "Buffalo",
          "state": "NY",
          "postalCode": "98052",
          "countryOrRegion": "USA"
        },
        "geoCoordinates": {
          "latitude": 47.6405,
          "longitude": -122.1293
        },
        "phone": "000-000-0000",
        "nickname": "Conf Room",
        "label": "200",
        "capacity": 40,
        "building": "2",
        "floorNumber": 2,
        "isManaged": true,
        "isWheelChairAccessible": false,
        "bookingType": "standard",
        "tags": [
          "benches",
          "nice view"
        ],
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDevice": "surface hub"
      }
    ]
  };
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
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROOM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'phone', 'emailAddress']);
  });

  it('lists all available rooms in the tenant (verbose)', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/places/microsoft.graph.room`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('lists all available rooms filter by roomlistEmail in the tenant (verbose)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/places/bldg2@contoso.com/microsoft.graph.roomlist/rooms`) {
          return jsonOutputFilter;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { verbose: true, roomlistEmail: "bldg2@contoso.com" } });
      assert(loggerLogSpy.calledWith(
        jsonOutputFilter.value
      ));
    }
  );

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: { force: true } }), new CommandError(errorMessage));
  });
});
