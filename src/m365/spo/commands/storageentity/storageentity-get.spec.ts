import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './storageentity-get.js';

describe(commands.STORAGEENTITY_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('existingproperty')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { Comment: 'Lorem', Description: 'ipsum', Value: 'dolor' };
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('propertywithoutdescription')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { Comment: 'Lorem', Value: 'dolor' };
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('propertywithoutcomments')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { Description: 'ipsum', Value: 'dolor' };
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('nonexistingproperty')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { "odata.null": true };
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('%23myprop')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { Description: 'ipsum', Value: 'dolor' };
        }
      }

      throw 'Invalid request';
    });
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.STORAGEENTITY_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the details of an existing tenant property', async () => {
    await command.action(logger, { options: { debug: true, key: 'existingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert(loggerLogSpy.calledWith({
      Key: 'existingproperty',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: 'Lorem'
    }));
  });

  it('retrieves the details of an existing tenant property without a description', async () => {
    await command.action(logger, { options: { debug: true, key: 'propertywithoutdescription', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert(loggerLogSpy.calledWith({
      Key: 'propertywithoutdescription',
      Value: 'dolor',
      Description: undefined,
      Comment: 'Lorem'
    }));
  });

  it('retrieves the details of an existing tenant property without a comment', async () => {
    await command.action(logger, { options: { key: 'propertywithoutcomments', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert(loggerLogSpy.calledWith({
      Key: 'propertywithoutcomments',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: undefined
    }));
  });

  it('handles a non-existent tenant property', async () => {
    await command.action(logger, { options: { key: 'nonexistingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
  });

  it('handles a non-existent tenant property (debug)', async () => {
    await command.action(logger, { options: { debug: true, key: 'nonexistingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    let correctValue: boolean = false;
    log.forEach(l => {
      if (l &&
        typeof l === 'string' &&
        l.indexOf('Property with key nonexistingproperty not found') > -1) {
        correctValue = true;
      }
    });
    assert(correctValue);
  });

  it('escapes special characters in property name', async () => {
    await command.action(logger, { options: { debug: true, key: '#myprop', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert(loggerLogSpy.calledWith({
      Key: '#myprop',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: undefined
    }));
  });

  it('requires tenant property name', () => {
    const options = command.options;
    let requiresTenantPropertyName = false;
    options.forEach(o => {
      if (o.option.indexOf('<key>') > -1) {
        requiresTenantPropertyName = true;
      }
    });
    assert(requiresTenantPropertyName);
  });

  it('handles promise rejection', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('error'));

    await assert.rejects(command.action(logger, { options: { debug: true, key: '#myprop' } } as any), new CommandError('error'));
  });
});
