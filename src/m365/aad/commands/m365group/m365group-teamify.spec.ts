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
import command from './m365group-teamify.js';
import { settingsNames } from '../../../../settingsNames.js';
import { aadGroup } from '../../../../utils/aadGroup.js';

describe(commands.M365GROUP_TEAMIFY, () => {
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
    jest.spyOn(aadGroup, 'isUnifiedGroup').mockClear().mockImplementation().resolves(true);
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
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.put,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_TEAMIFY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and mailNickname options are not passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('fails validation if both id and mailNickname options are passed',
    async () => {
      jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
        if (settingName === settingsNames.prompt) {
          return false;
        }

        return defaultValue;
      });

      const actual = await command.validate({
        options: {
          id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
          mailNickname: 'GroupName'
        }
      }, commandInfo);
      assert.notStrictEqual(actual, true);
    }
  );

  it('validates for a correct id', async () => {
    const actual = await command.validate({
      options: {
        id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to get M365 group when it does not exists', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=mailNickname eq '`) > -1) {
        return { value: [] };
      }
      throw 'The specified Microsoft 365 Group does not exist';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        mailNickname: 'GroupName'
      }
    }), new CommandError(`The specified Microsoft 365 Group does not exist`));
  });

  it('fails when multiple groups with same name exists', async () => {
    jest.spyOn(cli, 'getSettingWithDefaultValue').mockClear().mockImplementation((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=mailNickname eq '`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:01:19Z",
              "creationOptions": [],
              "description": "GroupName",
              "displayName": "GroupName",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "groupname@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "groupname",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [
                "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:groupname@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-09-05T09:01:19Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": []
            },
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:01:19Z",
              "creationOptions": [],
              "description": "GroupName",
              "displayName": "GroupName",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "groupname@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "groupname",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [
                "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:groupname@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-09-05T09:01:19Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": []
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        mailNickname: 'GroupName'
      }
    }), new CommandError("Multiple Microsoft 365 Groups with name 'GroupName' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === "https://graph.microsoft.com/v1.0/groups?$filter=mailNickname eq 'groupname'") {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
            "value": [
              {
                "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
                "id": "00000000-0000-0000-0000-000000000000",
                "deletedDateTime": null,
                "classification": null,
                "createdDateTime": "2021-09-05T09:01:19Z",
                "creationOptions": [],
                "description": "GroupName",
                "displayName": "GroupName",
                "expirationDateTime": null,
                "groupTypes": [
                  "Unified"
                ],
                "isAssignableToRole": null,
                "mail": "groupname@contoso.onmicrosoft.com",
                "mailEnabled": true,
                "mailNickname": "groupname",
                "membershipRule": null,
                "membershipRuleProcessingState": null,
                "onPremisesDomainName": null,
                "onPremisesLastSyncDateTime": null,
                "onPremisesNetBiosName": null,
                "onPremisesSamAccountName": null,
                "onPremisesSecurityIdentifier": null,
                "onPremisesSyncEnabled": null,
                "preferredDataLocation": null,
                "preferredLanguage": null,
                "proxyAddresses": [
                  "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                  "SMTP:groupname@contoso.onmicrosoft.com"
                ],
                "renewedDateTime": "2021-09-05T09:01:19Z",
                "resourceBehaviorOptions": [],
                "resourceProvisioningOptions": [
                  "Team"
                ],
                "securityEnabled": false,
                "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
                "theme": null,
                "visibility": "Public",
                "onPremisesProvisioningErrors": []
              },
              {
                "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
                "id": "00000000-0000-0000-0000-000000000000",
                "deletedDateTime": null,
                "classification": null,
                "createdDateTime": "2021-09-05T09:01:19Z",
                "creationOptions": [],
                "description": "GroupName",
                "displayName": "GroupName",
                "expirationDateTime": null,
                "groupTypes": [
                  "Unified"
                ],
                "isAssignableToRole": null,
                "mail": "groupname@contoso.onmicrosoft.com",
                "mailEnabled": true,
                "mailNickname": "groupname",
                "membershipRule": null,
                "membershipRuleProcessingState": null,
                "onPremisesDomainName": null,
                "onPremisesLastSyncDateTime": null,
                "onPremisesNetBiosName": null,
                "onPremisesSamAccountName": null,
                "onPremisesSecurityIdentifier": null,
                "onPremisesSyncEnabled": null,
                "preferredDataLocation": null,
                "preferredLanguage": null,
                "proxyAddresses": [
                  "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                  "SMTP:groupname@contoso.onmicrosoft.com"
                ],
                "renewedDateTime": "2021-09-05T09:01:19Z",
                "resourceBehaviorOptions": [],
                "resourceProvisioningOptions": [
                  "Team"
                ],
                "securityEnabled": false,
                "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
                "theme": null,
                "visibility": "Public",
                "onPremisesProvisioningErrors": []
              }
            ]
          };
        }
        throw 'Invalid request';
      });

      jest.spyOn(Cli, 'handleMultipleResultsFound').mockClear().mockImplementation().resolves({
        "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
        "id": "00000000-0000-0000-0000-000000000000",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2021-09-05T09:01:19Z",
        "creationOptions": [],
        "description": "GroupName",
        "displayName": "GroupName",
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "isAssignableToRole": null,
        "mail": "groupname@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "groupname",
        "membershipRule": null,
        "membershipRuleProcessingState": null,
        "onPremisesDomainName": null,
        "onPremisesLastSyncDateTime": null,
        "onPremisesNetBiosName": null,
        "onPremisesSamAccountName": null,
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "preferredLanguage": null,
        "proxyAddresses": [
          "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
          "SMTP:groupname@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2021-09-05T09:01:19Z",
        "resourceBehaviorOptions": [],
        "resourceProvisioningOptions": [
          "Team"
        ],
        "securityEnabled": false,
        "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
        "theme": null,
        "visibility": "Public",
        "onPremisesProvisioningErrors": []
      });

      const requestStub: jest.Mock = jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/team`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
            "id": "00000000-0000-0000-0000-000000000000",
            "createdDateTime": null,
            "displayName": "Group Team",
            "description": "Group Team description",
            "internalId": "19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01@thread.tacv2",
            "classification": null,
            "specialization": null,
            "mailNickname": "groupname",
            "visibility": "public",
            "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=00000000-0000-0000-0000-000000000000&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
            "isArchived": null,
            "isMembershipLimitedToOwners": false,
            "discoverySettings": null,
            "memberSettings": {
              "allowCreateUpdateChannels": true,
              "allowCreatePrivateChannels": true,
              "allowDeleteChannels": true,
              "allowAddRemoveApps": true,
              "allowCreateUpdateRemoveTabs": true,
              "allowCreateUpdateRemoveConnectors": true
            },
            "guestSettings": {
              "allowCreateUpdateChannels": false,
              "allowDeleteChannels": false
            },
            "messagingSettings": {
              "allowUserEditMessages": true,
              "allowUserDeleteMessages": true,
              "allowOwnerDeleteMessages": true,
              "allowTeamMentions": true,
              "allowChannelMentions": true
            },
            "funSettings": {
              "allowGiphy": true,
              "giphyContentRating": "moderate",
              "allowStickersAndMemes": true,
              "allowCustomMemes": true
            }
          };
        }
        throw 'Invalid request';
      });

      await command.action(logger, {
        options: { mailNickname: 'groupname' }
      });
      assert.strictEqual(requestStub.mock.lastCall[0].url, 'https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/team');
    }
  );

  it('Teamify M365 group by id', async () => {
    const requestStub: jest.Mock = jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/team`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": "8231f9f2-701f-4c6e-93ce-ecb563e3c1ee",
          "createdDateTime": null,
          "displayName": "Group Team",
          "description": "Group Team description",
          "internalId": "19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01@thread.tacv2",
          "classification": null,
          "specialization": null,
          "mailNickname": "groupname",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=8231f9f2-701f-4c6e-93ce-ecb563e3c1ee&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
          "isArchived": null,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": null,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowCreatePrivateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          },
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          },
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": true
          }
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    });
    assert.strictEqual(requestStub.mock.lastCall[0].url, 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/team');
  });

  it('Teamify M365 group by mailNickname', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=mailNickname eq `) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:01:19Z",
              "creationOptions": [],
              "description": "GroupName",
              "displayName": "GroupName",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "groupname@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "groupname",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [
                "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:groupname@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-09-05T09:01:19Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": []
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const requestStub: jest.Mock = jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/team`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": "00000000-0000-0000-0000-000000000000",
          "createdDateTime": null,
          "displayName": "Group Team",
          "description": "Group Team description",
          "internalId": "19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01@thread.tacv2",
          "classification": null,
          "specialization": null,
          "mailNickname": "groupname",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=00000000-0000-0000-0000-000000000000&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
          "isArchived": null,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": null,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowCreatePrivateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          },
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          },
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": true
          }
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { mailNickname: 'groupname' }
    });
    assert.strictEqual(requestStub.mock.lastCall[0].url, 'https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/team');
  });

  it('should handle Microsoft graph error response', async () => {
    jest.spyOn(request, 'put').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/team`) {
        throw {
          "error": {
            "code": "NotFound",
            "message": "Failed to execute MS Graph backend request GetGroupInternalApiRequest",
            "innerError": {
              "date": "2021-06-19T03:00:13",
              "request-id": "0e3f93f6-d3f7-4d84-9eb5-dc2dda0eec0e",
              "client-request-id": "68cff2aa-b010-daa7-2467-fa8e96cbda25"
            }
          }
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    } as any), new CommandError('Failed to execute MS Graph backend request GetGroupInternalApiRequest'));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    jestUtil.restore(aadGroup.isUnifiedGroup);
    jest.spyOn(aadGroup, 'isUnifiedGroup').mockClear().mockImplementation().resolves(false);

    await assert.rejects(command.action(logger, { options: { id: groupId } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
