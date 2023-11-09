import assert from 'assert';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { AuthType, CertificateType, CloudType, Service } from '../Auth.js';
import { jestUtil } from '../utils/jestUtil.js';
import { FileTokenStorage } from './FileTokenStorage.js';

describe('FileTokenStorage', () => {
  const fileStorage = new FileTokenStorage(FileTokenStorage.connectionInfoFilePath());

  afterEach(() => {
    jestUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFile
    ]);
  });

  it(`stores MSAL cache in the user's home directory`, () => {
    assert.strictEqual(FileTokenStorage.msalCacheFilePath(), path.join(os.homedir(), '.cli-m365-msal.json'));
  });

  it('fails retrieving connection info from file if the token file doesn\'t exist',
    (done) => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      fileStorage
        .get()
        .then(() => {
          done('Expected fail but passed instead');
        }, (err) => {
          try {
            assert.strictEqual(err, 'File not found');
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );

  it('returns connection info from file', (done) => {
    const tokensFile: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => JSON.stringify(tokensFile));
    fileStorage
      .get()
      .then((connectionInfo) => {
        try {
          assert.strictEqual(connectionInfo, JSON.stringify(tokensFile));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the connection info in the file when the file doesn\'t exist',
    (done) => {
      const expected: Service = {
        accessTokens: {},
        appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        tenant: 'common',
        cloudType: CloudType.Public,
        authType: AuthType.DeviceCode,
        certificateType: CertificateType.Unknown,
        connected: false,
        logout: () => { }
      };
      let actual: string = '';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation((path, token) => { actual = token as string; }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .set(JSON.stringify(expected))
        .then(() => {
          try {
            assert.strictEqual(actual, JSON.stringify(expected));
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );

  it('saves the connection info in the file when the file is empty',
    (done) => {
      const expected: Service = {
        accessTokens: {},
        appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        tenant: 'common',
        cloudType: CloudType.Public,
        authType: AuthType.DeviceCode,
        certificateType: CertificateType.Unknown,
        connected: false,
        logout: () => { }
      };
      let actual: string = '';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => '');
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation((path, token) => { actual = token as string; }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .set(JSON.stringify(expected))
        .then(() => {
          try {
            assert.strictEqual(actual, JSON.stringify(expected));
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );

  it('saves the connection info in the file when the file contains an empty JSON object',
    (done) => {
      const expected: Service = {
        accessTokens: {},
        appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        tenant: 'common',
        cloudType: CloudType.Public,
        authType: AuthType.DeviceCode,
        certificateType: CertificateType.Unknown,
        connected: false,
        logout: () => { }
      };
      let actual: string = '';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => '{}');
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation((path, token) => { actual = token as string; }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .set(JSON.stringify(expected))
        .then(() => {
          try {
            assert.strictEqual(actual, JSON.stringify(expected));
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );

  it('saves the connection info in the file when the file contains no access tokens',
    (done) => {
      const expected: Service = {
        accessTokens: {},
        appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        tenant: 'common',
        cloudType: CloudType.Public,
        authType: AuthType.DeviceCode,
        certificateType: CertificateType.Unknown,
        connected: false,
        logout: () => { }
      };
      let actual: string = '';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => '{"accessTokens":{},"authType":0,"connected":false}');
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation((path, token) => { actual = token as string; }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .set(JSON.stringify(expected))
        .then(() => {
          try {
            assert.strictEqual(actual, JSON.stringify(expected));
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );

  it('adds the connection info to the file when the file contains access tokens',
    (done) => {
      const expected: Service = {
        accessTokens: {},
        appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
        tenant: 'common',
        cloudType: CloudType.Public,
        authType: AuthType.DeviceCode,
        certificateType: CertificateType.Unknown,
        connected: false,
        logout: () => { }
      };
      let actual: string = '';
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => JSON.stringify({
        accessTokens: {
          "https://contoso.sharepoint.com": {
            expiresOn: '123',
            value: '123'
          }
        },
        authType: AuthType.DeviceCode,
        connected: true,
        refreshToken: 'ref'
      }));
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation((path, token) => { actual = token as string; }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .set(JSON.stringify(expected))
        .then(() => {
          try {
            assert.strictEqual(actual, JSON.stringify(expected));
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );

  it('correctly handles error when writing to the file failed', (done) => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    jest.spyOn(fs, 'writeFile').mockClear().mockImplementation(() => { }).mockImplementation((...args: any[]) => args[3]({ message: 'An error has occurred' }));
    fileStorage
      .set('abc')
      .then(() => {
        done('Fail expected but passed instead');
      }, (err) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('succeeds with removing if the token file doesn\'t exist', (done) => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    fileStorage
      .remove()
      .then(() => {
        done();
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file is empty', (done) => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => '');
    jest.spyOn(fs, 'writeFile').mockClear().mockImplementation(() => '').mockImplementation((...args: any[]) => args[3](null));
    fileStorage
      .remove()
      .then(() => {
        done();
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file contains empty JSON object',
    (done) => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => '{}');
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation(() => '').mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .remove()
        .then(() => {
          done();
        }, () => {
          done('Pass expected but failed instead');
        });
    }
  );

  it('succeeds with removing if the token file contains no services',
    (done) => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => JSON.stringify({ services: {} }));
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation(() => { }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .remove()
        .then(() => {
          done();
        }, () => {
          done('Pass expected but failed instead');
        });
    }
  );

  it('succeeds when connection info successfully removed from the token file',
    (done) => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => JSON.stringify({
        services: {
          'abc': 'def'
        }
      }));
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation(() => { }).mockImplementation((...args: any[]) => args[3](null));
      fileStorage
        .remove()
        .then(() => {
          try {
            done();
          }
          catch (e) {
            done(e);
          }
        }, () => {
          done('Pass expected but failed instead');
        });
    }
  );

  it('correctly handles error when writing updated tokens to the token file',
    (done) => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => true);
      jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation(() => JSON.stringify({
        services: {
          'abc': 'def'
        }
      }));
      jest.spyOn(fs, 'writeFile').mockClear().mockImplementation(() => { }).mockImplementation((...args: any[]) => args[3]({ message: 'An error has occurred' }));
      fileStorage
        .remove()
        .then(() => {
          done('Fail expected but passed instead');
        }, (err) => {
          try {
            assert.strictEqual(err, 'An error has occurred');
            done();
          }
          catch (e) {
            done(e);
          }
        });
    }
  );
});