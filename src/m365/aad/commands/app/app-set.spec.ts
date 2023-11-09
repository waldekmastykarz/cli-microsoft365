import assert from 'assert';
import fs from 'fs';
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
import command from './app-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_SET, () => {

  //#region Mocked Responses  
  const appDetailsResponse: any = { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity", "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83", "deletedDateTime": null, "appId": "ff254847-12c7-44cf-921e-8883dbd622a7", "applicationTemplateId": null, "disabledByMicrosoftStatus": null, "createdDateTime": "2022-02-07T08:51:18Z", "displayName": "Angular Teams app", "description": null, "groupMembershipClaims": null, "identifierUris": ["api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"], "isDeviceOnlyAuthSupported": null, "isFallbackPublicClient": null, "notes": null, "publisherDomain": "contoso.onmicrosoft.com", "serviceManagementReference": null, "signInAudience": "AzureADMyOrg", "tags": [], "tokenEncryptionKeyId": null, "defaultRedirectUri": null, "certification": null, "optionalClaims": null, "addIns": [], "api": { "acceptMappedClaims": null, "knownClientApplications": [], "requestedAccessTokenVersion": null, "oauth2PermissionScopes": [{ "adminConsentDescription": "Access as a user", "adminConsentDisplayName": "Access as a user", "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5", "isEnabled": true, "type": "User", "userConsentDescription": null, "userConsentDisplayName": null, "value": "access_as_user" }], "preAuthorizedApplications": [{ "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346", "delegatedPermissionIds": ["cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"] }, { "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264", "delegatedPermissionIds": ["cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"] }] }, "appRoles": [], "info": { "logoUrl": null, "marketingUrl": null, "privacyStatementUrl": null, "supportUrl": null, "termsOfServiceUrl": null }, "keyCredentials": [], "parentalControlSettings": { "countriesBlockedForMinors": [], "legalAgeGroupRule": "Allow" }, "passwordCredentials": [], "publicClient": { "redirectUris": [] }, "requiredResourceAccess": [{ "resourceAppId": "00000003-0000-0000-c000-000000000000", "resourceAccess": [{ "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope" }] }], "verifiedPublisher": { "displayName": null, "verifiedPublisherId": null, "addedDateTime": null }, "web": { "homePageUrl": null, "logoutUrl": null, "redirectUris": [], "implicitGrantSettings": { "enableAccessTokenIssuance": false, "enableIdTokenIssuance": false } }, "spa": { "redirectUris": [] } };
  //#endregion

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
      request.get,
      request.patch,
      fs.existsSync,
      fs.readFileSync,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates uris for the specified appId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq 'bc724b77-da87-43a9-b385-6ebaaf969db8'&$select=id`) {
        return {
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    });
  });

  it('updates multiple URIs for the specified appId', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq 'bc724b77-da87-43a9-b385-6ebaaf969db8'&$select=id`) {
        return {
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' &&
        opts.data.identifierUris[1] === 'api://testapi') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appId: 'bc724b77-da87-43a9-b385-6ebaaf969db8',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8,api://testapi'
      }
    });
  });

  it('updates uris for the specified objectId', async () => {
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    });
  });

  it('updates multiple URIs for the specified objectId', async () => {
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' &&
        opts.data.identifierUris[1] === 'api://testapi') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8,api://testapi'
      }
    });
  });

  it('updates uris for the specified name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return {
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    });
  });

  it('updates multiple URIs for the specified name', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return {
          value: [{
            id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
          }]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.identifierUris[0] === 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' &&
        opts.data.identifierUris[1] === 'api://testapi') {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8,api://testapi'
      }
    });
  });

  it('skips updating uris if no uris specified', async () => {
    await command.action(logger, {
      options: {
        objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230'
      }
    });
  });

  it('sets spa redirectUri for an app without redirectUris', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/e4528262-097a-42eb-98e1-19f073dbee45`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "e4528262-097a-42eb-98e1-19f073dbee45",
          "deletedDateTime": null,
          "appId": "842e1a6f-7492-4b7d-8278-563036f5bd39",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-10T08:01:06Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5196.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": []
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/e4528262-097a-42eb-98e1-19f073dbee45' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth"
            ]
          },
          "web": {
            "redirectUris": []
          }
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        objectId: 'e4528262-097a-42eb-98e1-19f073dbee45',
        redirectUris: 'https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
        platform: 'spa'
      }
    });
  });

  it('sets web redirectUri for an app with existing spa redirectUris',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
            "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
            "deletedDateTime": null,
            "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
            "applicationTemplateId": null,
            "disabledByMicrosoftStatus": null,
            "createdDateTime": "2022-02-07T08:51:18Z",
            "displayName": "Angular Teams app",
            "description": null,
            "groupMembershipClaims": null,
            "identifierUris": [
              "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
            ],
            "isDeviceOnlyAuthSupported": null,
            "isFallbackPublicClient": null,
            "notes": null,
            "publisherDomain": "contoso.onmicrosoft.com",
            "serviceManagementReference": null,
            "signInAudience": "AzureADMyOrg",
            "tags": [],
            "tokenEncryptionKeyId": null,
            "defaultRedirectUri": null,
            "certification": null,
            "optionalClaims": null,
            "addIns": [],
            "api": {
              "acceptMappedClaims": null,
              "knownClientApplications": [],
              "requestedAccessTokenVersion": null,
              "oauth2PermissionScopes": [
                {
                  "adminConsentDescription": "Access as a user",
                  "adminConsentDisplayName": "Access as a user",
                  "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                  "isEnabled": true,
                  "type": "User",
                  "userConsentDescription": null,
                  "userConsentDisplayName": null,
                  "value": "access_as_user"
                }
              ],
              "preAuthorizedApplications": [
                {
                  "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                  "delegatedPermissionIds": [
                    "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                  ]
                },
                {
                  "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                  "delegatedPermissionIds": [
                    "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                  ]
                }
              ]
            },
            "appRoles": [],
            "info": {
              "logoUrl": null,
              "marketingUrl": null,
              "privacyStatementUrl": null,
              "supportUrl": null,
              "termsOfServiceUrl": null
            },
            "keyCredentials": [],
            "parentalControlSettings": {
              "countriesBlockedForMinors": [],
              "legalAgeGroupRule": "Allow"
            },
            "passwordCredentials": [],
            "publicClient": {
              "redirectUris": []
            },
            "requiredResourceAccess": [
              {
                "resourceAppId": "00000003-0000-0000-c000-000000000000",
                "resourceAccess": [
                  {
                    "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                    "type": "Scope"
                  }
                ]
              }
            ],
            "verifiedPublisher": {
              "displayName": null,
              "verifiedPublisherId": null,
              "addedDateTime": null
            },
            "web": {
              "homePageUrl": null,
              "logoutUrl": null,
              "redirectUris": [],
              "implicitGrantSettings": {
                "enableAccessTokenIssuance": false,
                "enableIdTokenIssuance": false
              }
            },
            "spa": {
              "redirectUris": [
                "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
                "http://localhost/auth"
              ]
            }
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
          JSON.stringify(opts.data) === JSON.stringify({
            "publicClient": {
              "redirectUris": []
            },
            "spa": {
              "redirectUris": [
                "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
                "http://localhost/auth"
              ]
            },
            "web": {
              "redirectUris": [
                "https://foo.com"
              ]
            }
          })) {
          return;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await command.action(logger, {
        options: {
          debug: true,
          objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
          redirectUris: 'https://foo.com',
          platform: 'web'
        }
      });
    }
  );

  it('sets publicClient redirectUri for an app with existing spa and web redirectUris',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
            "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
            "deletedDateTime": null,
            "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
            "applicationTemplateId": null,
            "disabledByMicrosoftStatus": null,
            "createdDateTime": "2022-02-07T08:51:18Z",
            "displayName": "Angular Teams app",
            "description": null,
            "groupMembershipClaims": null,
            "identifierUris": [
              "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
            ],
            "isDeviceOnlyAuthSupported": null,
            "isFallbackPublicClient": null,
            "notes": null,
            "publisherDomain": "contoso.onmicrosoft.com",
            "serviceManagementReference": null,
            "signInAudience": "AzureADMyOrg",
            "tags": [],
            "tokenEncryptionKeyId": null,
            "defaultRedirectUri": null,
            "certification": null,
            "optionalClaims": null,
            "addIns": [],
            "api": {
              "acceptMappedClaims": null,
              "knownClientApplications": [],
              "requestedAccessTokenVersion": null,
              "oauth2PermissionScopes": [
                {
                  "adminConsentDescription": "Access as a user",
                  "adminConsentDisplayName": "Access as a user",
                  "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                  "isEnabled": true,
                  "type": "User",
                  "userConsentDescription": null,
                  "userConsentDisplayName": null,
                  "value": "access_as_user"
                }
              ],
              "preAuthorizedApplications": [
                {
                  "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                  "delegatedPermissionIds": [
                    "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                  ]
                },
                {
                  "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                  "delegatedPermissionIds": [
                    "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                  ]
                }
              ]
            },
            "appRoles": [],
            "info": {
              "logoUrl": null,
              "marketingUrl": null,
              "privacyStatementUrl": null,
              "supportUrl": null,
              "termsOfServiceUrl": null
            },
            "keyCredentials": [],
            "parentalControlSettings": {
              "countriesBlockedForMinors": [],
              "legalAgeGroupRule": "Allow"
            },
            "passwordCredentials": [],
            "publicClient": {
              "redirectUris": []
            },
            "requiredResourceAccess": [
              {
                "resourceAppId": "00000003-0000-0000-c000-000000000000",
                "resourceAccess": [
                  {
                    "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                    "type": "Scope"
                  }
                ]
              }
            ],
            "verifiedPublisher": {
              "displayName": null,
              "verifiedPublisherId": null,
              "addedDateTime": null
            },
            "web": {
              "homePageUrl": null,
              "logoutUrl": null,
              "redirectUris": [
                "https://foo.com"
              ],
              "implicitGrantSettings": {
                "enableAccessTokenIssuance": false,
                "enableIdTokenIssuance": false
              }
            },
            "spa": {
              "redirectUris": [
                "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
                "http://localhost/auth"
              ]
            }
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
          JSON.stringify(opts.data) === JSON.stringify({
            "publicClient": {
              "redirectUris": [
                "https://foo1.com"
              ]
            },
            "spa": {
              "redirectUris": [
                "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
                "http://localhost/auth"
              ]
            },
            "web": {
              "redirectUris": [
                "https://foo.com"
              ]
            }
          })) {
          return;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await command.action(logger, {
        options: {
          debug: true,
          objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
          redirectUris: 'https://foo1.com',
          platform: 'publicClient'
        }
      });
    }
  );

  it('replaces existing redirectUri with a new one', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
          "deletedDateTime": null,
          "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
          "applicationTemplateId": null,
          "disabledByMicrosoftStatus": null,
          "createdDateTime": "2022-02-07T08:51:18Z",
          "displayName": "Angular Teams app",
          "description": null,
          "groupMembershipClaims": null,
          "identifierUris": [
            "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
          ],
          "isDeviceOnlyAuthSupported": null,
          "isFallbackPublicClient": null,
          "notes": null,
          "publisherDomain": "contoso.onmicrosoft.com",
          "serviceManagementReference": null,
          "signInAudience": "AzureADMyOrg",
          "tags": [],
          "tokenEncryptionKeyId": null,
          "defaultRedirectUri": null,
          "certification": null,
          "optionalClaims": null,
          "addIns": [],
          "api": {
            "acceptMappedClaims": null,
            "knownClientApplications": [],
            "requestedAccessTokenVersion": null,
            "oauth2PermissionScopes": [
              {
                "adminConsentDescription": "Access as a user",
                "adminConsentDisplayName": "Access as a user",
                "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                "isEnabled": true,
                "type": "User",
                "userConsentDescription": null,
                "userConsentDisplayName": null,
                "value": "access_as_user"
              }
            ],
            "preAuthorizedApplications": [
              {
                "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              },
              {
                "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                "delegatedPermissionIds": [
                  "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                ]
              }
            ]
          },
          "appRoles": [],
          "info": {
            "logoUrl": null,
            "marketingUrl": null,
            "privacyStatementUrl": null,
            "supportUrl": null,
            "termsOfServiceUrl": null
          },
          "keyCredentials": [],
          "parentalControlSettings": {
            "countriesBlockedForMinors": [],
            "legalAgeGroupRule": "Allow"
          },
          "passwordCredentials": [],
          "publicClient": {
            "redirectUris": []
          },
          "requiredResourceAccess": [
            {
              "resourceAppId": "00000003-0000-0000-c000-000000000000",
              "resourceAccess": [
                {
                  "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                  "type": "Scope"
                }
              ]
            }
          ],
          "verifiedPublisher": {
            "displayName": null,
            "verifiedPublisherId": null,
            "addedDateTime": null
          },
          "web": {
            "homePageUrl": null,
            "logoutUrl": null,
            "redirectUris": [],
            "implicitGrantSettings": {
              "enableAccessTokenIssuance": false,
              "enableIdTokenIssuance": false
            }
          },
          "spa": {
            "redirectUris": [
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
              "http://localhost/auth"
            ]
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "publicClient": {
            "redirectUris": []
          },
          "spa": {
            "redirectUris": [
              "http://localhost/auth",
              "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth"
            ]
          },
          "web": {
            "redirectUris": []
          }
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        redirectUris: 'https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
        platform: 'spa',
        redirectUrisToRemove: 'https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth'
      }
    });
  });

  it('adds new certificate using base64 string', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateBase64Encoded: 'somecertificatebase64string'
      }
    });
  });

  it('adds new certificate using base64 string (with null keyCredentials)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
          return appDetailsResponse;
        }
        else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
            "keyCredentials": null
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
          JSON.stringify(opts.data) === JSON.stringify({
            "keyCredentials": [{
              "type": "AsymmetricX509Cert",
              "usage": "Verify",
              "displayName": "some certificate",
              "key": "somecertificatebase64string"
            }]
          })) {
          return;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await command.action(logger, {
        options: {
          debug: true,
          objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
          certificateDisplayName: 'some certificate',
          certificateBase64Encoded: 'somecertificatebase64string'
        }
      });
    }
  );

  it('adds new certificate using certificate file', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83?$select=keyCredentials`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
          "keyCredentials": []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "keyCredentials": [{
            "type": "AsymmetricX509Cert",
            "usage": "Verify",
            "displayName": "some certificate",
            "key": "somecertificatebase64string"
          }]
        })) {
        return;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockReturnValue("somecertificatebase64string");

    await command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateFile: 'C:\\temp\\some-certificate.cer'
      }
    });
  });

  it('handles error when certificate file cannot be read', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/95cfe30d-ed44-4f9d-b73d-c66560f72e83`) {
        return appDetailsResponse;
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(fs, 'existsSync').mockClear().mockReturnValue(true);
    jest.spyOn(fs, 'readFileSync').mockClear().mockImplementation().throws(new Error("An error has occurred"));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83',
        certificateDisplayName: 'some certificate',
        certificateFile: 'C:\\temp\\some-certificate.cer'
      }
    }), new CommandError(`Error reading certificate file: Error: An error has occurred. Please add the certificate using base64 option '--certificateBase64Encoded'.`));
  });

  it('handles error when the app specified with objectId not found',
    async () => {
      jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects(new Error(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));

      await assert.rejects(command.action(logger, {
        options: {
          objectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
        }
      }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
    }
  );

  it('handles error when the app specified with the appId not found',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
          return { value: [] };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects('PATCH request executed');

      await assert.rejects(command.action(logger, {
        options: {
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
        }
      }), new CommandError(`No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
    }
  );

  it('handles error when the app specified with name not found', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });
    jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects('PATCH request executed');

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My app',
        uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
      }
    }), new CommandError(`No Azure AD application registration with name My app found`));
  });

  it('handles error when multiple apps with the specified name found',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });
      jest.spyOn(request, 'patch').mockClear().mockImplementation().rejects('PATCH request executed');

      await assert.rejects(command.action(logger, {
        options: {
          name: 'My app',
          uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
        }
      }), new CommandError(`Multiple Azure AD application registration with name 'My app' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
    }
  );

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt',
    async () => {
      let updateRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async opts => {
        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
          return {
            value: [
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
              { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
            ]
          };
        }

        if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#applications/$entity",
            "id": "95cfe30d-ed44-4f9d-b73d-c66560f72e83",
            "deletedDateTime": null,
            "appId": "ff254847-12c7-44cf-921e-8883dbd622a7",
            "applicationTemplateId": null,
            "disabledByMicrosoftStatus": null,
            "createdDateTime": "2022-02-07T08:51:18Z",
            "displayName": "Angular Teams app",
            "description": null,
            "groupMembershipClaims": null,
            "identifierUris": [
              "api://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/ff254847-12c7-44cf-921e-8883dbd622a7"
            ],
            "isDeviceOnlyAuthSupported": null,
            "isFallbackPublicClient": null,
            "notes": null,
            "publisherDomain": "contoso.onmicrosoft.com",
            "serviceManagementReference": null,
            "signInAudience": "AzureADMyOrg",
            "tags": [],
            "tokenEncryptionKeyId": null,
            "defaultRedirectUri": null,
            "certification": null,
            "optionalClaims": null,
            "addIns": [],
            "api": {
              "acceptMappedClaims": null,
              "knownClientApplications": [],
              "requestedAccessTokenVersion": null,
              "oauth2PermissionScopes": [
                {
                  "adminConsentDescription": "Access as a user",
                  "adminConsentDisplayName": "Access as a user",
                  "id": "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5",
                  "isEnabled": true,
                  "type": "User",
                  "userConsentDescription": null,
                  "userConsentDisplayName": null,
                  "value": "access_as_user"
                }
              ],
              "preAuthorizedApplications": [
                {
                  "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
                  "delegatedPermissionIds": [
                    "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                  ]
                },
                {
                  "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
                  "delegatedPermissionIds": [
                    "cf38eb5b-8fcd-4697-9bd5-d80b7f98dfc5"
                  ]
                }
              ]
            },
            "appRoles": [],
            "info": {
              "logoUrl": null,
              "marketingUrl": null,
              "privacyStatementUrl": null,
              "supportUrl": null,
              "termsOfServiceUrl": null
            },
            "keyCredentials": [],
            "parentalControlSettings": {
              "countriesBlockedForMinors": [],
              "legalAgeGroupRule": "Allow"
            },
            "passwordCredentials": [],
            "publicClient": {
              "redirectUris": []
            },
            "requiredResourceAccess": [
              {
                "resourceAppId": "00000003-0000-0000-c000-000000000000",
                "resourceAccess": [
                  {
                    "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
                    "type": "Scope"
                  }
                ]
              }
            ],
            "verifiedPublisher": {
              "displayName": null,
              "verifiedPublisherId": null,
              "addedDateTime": null
            },
            "web": {
              "homePageUrl": null,
              "logoutUrl": null,
              "redirectUris": [],
              "implicitGrantSettings": {
                "enableAccessTokenIssuance": false,
                "enableIdTokenIssuance": false
              }
            },
            "spa": {
              "redirectUris": [
                "https://244e-2001-1c00-80c-d00-e5da-977c-7c52-5193.ngrok.io/auth",
                "http://localhost/auth"
              ]
            }
          };
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({ id: '5b31c38c-2584-42f0-aa47-657fb3a84230' });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async opts => {
        if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230') {
          updateRequestIssued = true;
          return;
        }

        throw `Invalid request ${JSON.stringify(opts)}`;
      });

      await command.action(logger, {
        options: {
          debug: true,
          name: 'My app',
          redirectUris: 'https://24c4-2001-1c00-80c-d00-e5da-977c-7c52-5194.ngrok.io/auth',
          platform: 'spa'
        }
      });
      assert(updateRequestIssued);
    }
  );

  it('handles error when retrieving information about app through appId failed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

      await assert.rejects(command.action(logger, {
        options: {
          appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
          uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
        }
      }), new CommandError(`An error has occurred`));
    }
  );

  it('handles error when retrieving information about app through name failed',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().rejects(new Error('An error has occurred'));

      await assert.rejects(command.action(logger, {
        options: {
          name: 'My app',
          uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8'
        }
      }), new CommandError(`An error has occurred`));
    }
  );

  it('fails validation if appId and objectId specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and name specified', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if objectId and name specified', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, objectId nor name specified',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({ options: {} }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if redirectUris specified without platform',
    async () => {
      const actual = await command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', redirectUris: 'https://foo.com' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if invalid platform specified', async () => {
    const actual = await command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', redirectUris: 'https://foo.com', platform: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if certificateDisplayName is specified without certificate',
    async () => {
      const actual = await command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', certificateDisplayName: 'Some certificate' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both certificateBase64Encoded and certificateFile are specified',
    async () => {
      const actual = await command.validate({ options: { objectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8', certificateFile: 'c:\\temp\\some-certificate.cer', certificateBase64Encoded: 'somebase64string' } }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('passes validation if certificateFile specified with certificateDisplayName',
    async () => {
      jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => true);

      const actual = await command.validate({ options: { name: 'My AAD app', certificateDisplayName: 'Some certificate', certificateFile: 'c:\\temp\\some-certificate.cer' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('fails validation when certificate file is not found', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(_ => false);

    const actual = await command.validate({ options: { debug: true, objectId: '95cfe30d-ed44-4f9d-b73d-c66560f72e83', certificateDisplayName: 'some certificate', certificateFile: 'C:\\temp\\some-certificate.cer' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (appId)', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (objectId)',
    async () => {
      const actual = await command.validate({ options: { objectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { name: 'My app', uris: 'https://contoso.com/bc724b77-da87-43a9-b385-6ebaaf969db8' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when redirectUris specified with spa', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'spa' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when redirectUris specified with publicClient',
    async () => {
      const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'publicClient' } }, commandInfo);
      assert.strictEqual(actual, true);
    }
  );

  it('passes validation when redirectUris specified with web', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', redirectUris: 'https://foo.com', platform: 'web' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
