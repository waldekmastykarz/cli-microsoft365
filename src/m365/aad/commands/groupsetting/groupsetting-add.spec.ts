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
import command from './groupsetting-add.js';

describe(commands.GROUPSETTING_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: jest.SpyInstance;
  let commandInfo: CommandInfo;

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
    loggerLogSpy = jest.spyOn(logger, 'log').mockClear();
    (command as any).items = [];
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.post
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUPSETTING_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds group setting using default template setting values', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return {
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        };
      }

      throw 'Invalid Request';
    });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === JSON.stringify({
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'CustomBlockedWordsList',
              value: ''
            },
            {
              name: 'EnableMSStandardBlockedWords',
              value: 'false'
            },
            {
              name: 'ClassificationDescriptions',
              value: ''
            },
            {
              name: 'DefaultClassification',
              value: ''
            },
            {
              name: 'PrefixSuffixNamingRequirement',
              value: ''
            },
            {
              name: 'AllowGuestsToBeGroupOwner',
              value: 'false'
            },
            {
              name: 'AllowGuestsToAccessGroups',
              value: 'true'
            },
            {
              name: 'GuestUsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'GroupCreationAllowedGroupId',
              value: ''
            },
            {
              name: 'AllowToAddGuests',
              value: 'true'
            },
            {
              name: 'UsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'ClassificationList',
              value: ''
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return {
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b' } });
    assert(loggerLogSpy.calledWith({
      displayName: null,
      id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
      templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
      values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
    }));
  });

  it('adds group setting using default template setting values (debug)',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
          return {
            "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
          };
        }

        throw 'Invalid Request';
      });
      jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
          JSON.stringify(opts.data) === JSON.stringify({
            templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
            values: [
              {
                name: 'CustomBlockedWordsList',
                value: ''
              },
              {
                name: 'EnableMSStandardBlockedWords',
                value: 'false'
              },
              {
                name: 'ClassificationDescriptions',
                value: ''
              },
              {
                name: 'DefaultClassification',
                value: ''
              },
              {
                name: 'PrefixSuffixNamingRequirement',
                value: ''
              },
              {
                name: 'AllowGuestsToBeGroupOwner',
                value: 'false'
              },
              {
                name: 'AllowGuestsToAccessGroups',
                value: 'true'
              },
              {
                name: 'GuestUsageGuidelinesUrl',
                value: ''
              },
              {
                name: 'GroupCreationAllowedGroupId',
                value: ''
              },
              {
                name: 'AllowToAddGuests',
                value: 'true'
              },
              {
                name: 'UsageGuidelinesUrl',
                value: ''
              },
              {
                name: 'ClassificationList',
                value: ''
              },
              {
                name: 'EnableGroupCreation',
                value: 'true'
              }
            ]
          })) {
          return {
            displayName: null,
            id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
            templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
            values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
          };
        }

        throw 'Invalid Request';
      });

      await command.action(logger, { options: { debug: true, templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b' } });
      assert(loggerLogSpy.calledWith({
        displayName: null,
        id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
        templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
        values: [{ "name": "UsageGuidelinesUrl", "value": "" }, { "name": "ClassificationList", "value": "" }, { "name": "DefaultClassification", "value": "" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
      }));
    }
  );

  it('adds group setting using the specified values', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return {
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        };
      }

      throw 'Invalid Request';
    });
    jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === JSON.stringify({
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'UsageGuidelinesUrl',
              value: 'https://contoso.sharepoint.com/sites/compliance'
            },
            {
              name: 'ClassificationList',
              value: 'HBI, MBI, LBI, GDPR'
            },
            {
              name: 'DefaultClassification',
              value: 'MBI'
            },
            {
              name: 'CustomBlockedWordsList',
              value: ''
            },
            {
              name: 'EnableMSStandardBlockedWords',
              value: 'false'
            },
            {
              name: 'ClassificationDescriptions',
              value: ''
            },
            {
              name: 'PrefixSuffixNamingRequirement',
              value: ''
            },
            {
              name: 'AllowGuestsToBeGroupOwner',
              value: 'false'
            },
            {
              name: 'AllowGuestsToAccessGroups',
              value: 'true'
            },
            {
              name: 'GuestUsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'GroupCreationAllowedGroupId',
              value: ''
            },
            {
              name: 'AllowToAddGuests',
              value: 'true'
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return {
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b', UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance', ClassificationList: 'HBI, MBI, LBI, GDPR', DefaultClassification: 'MBI' } });
    assert(loggerLogSpy.calledWith({
      displayName: null,
      id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
      templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
      values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
    }));
  });

  it('ignores global options when creating request data', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return {
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        };
      }

      throw 'Invalid Request';
    });
    const postStub = jest.spyOn(request, 'post').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings` &&
        JSON.stringify(opts.data) === JSON.stringify({
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [
            {
              name: 'UsageGuidelinesUrl',
              value: 'https://contoso.sharepoint.com/sites/compliance'
            },
            {
              name: 'ClassificationList',
              value: 'HBI, MBI, LBI, GDPR'
            },
            {
              name: 'DefaultClassification',
              value: 'MBI'
            },
            {
              name: 'CustomBlockedWordsList',
              value: ''
            },
            {
              name: 'EnableMSStandardBlockedWords',
              value: 'false'
            },
            {
              name: 'ClassificationDescriptions',
              value: ''
            },
            {
              name: 'PrefixSuffixNamingRequirement',
              value: ''
            },
            {
              name: 'AllowGuestsToBeGroupOwner',
              value: 'false'
            },
            {
              name: 'AllowGuestsToAccessGroups',
              value: 'true'
            },
            {
              name: 'GuestUsageGuidelinesUrl',
              value: ''
            },
            {
              name: 'GroupCreationAllowedGroupId',
              value: ''
            },
            {
              name: 'AllowToAddGuests',
              value: 'true'
            },
            {
              name: 'EnableGroupCreation',
              value: 'true'
            }
          ]
        })) {
        return {
          displayName: null,
          id: 'cb9ede6b-fa00-474c-b34f-dae81102d210',
          templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
          values: [{ "name": "UsageGuidelinesUrl", "value": "https://contoso.sharepoint.com/sites/compliance" }, { "name": "ClassificationList", "value": "HBI, MBI, LBI, GDPR" }, { "name": "DefaultClassification", "value": "MBI" }, { "name": "CustomBlockedWordsList", "value": "" }, { "name": "EnableMSStandardBlockedWords", "value": "false" }, { "name": "ClassificationDescriptions", "value": "" }, { "name": "PrefixSuffixNamingRequirement", "value": "" }, { "name": "AllowGuestsToBeGroupOwner", "value": "false" }, { "name": "AllowGuestsToAccessGroups", "value": "true" }, { "name": "GuestUsageGuidelinesUrl", "value": "" }, { "name": "GroupCreationAllowedGroupId", "value": "" }, { "name": "AllowToAddGuests", "value": "true" }, { "name": "EnableGroupCreation", "value": "true" }]
        };
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, output: "text", templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b', UsageGuidelinesUrl: 'https://contoso.sharepoint.com/sites/compliance', ClassificationList: 'HBI, MBI, LBI, GDPR', DefaultClassification: 'MBI' } });
    assert.deepEqual(postStub.mock.calls[0][0].data, {
      templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b',
      values: [
        {
          name: 'UsageGuidelinesUrl',
          value: 'https://contoso.sharepoint.com/sites/compliance'
        },
        { name: 'ClassificationList', value: 'HBI, MBI, LBI, GDPR' },
        { name: 'DefaultClassification', value: 'MBI' },
        { name: 'CustomBlockedWordsList', value: '' },
        { name: 'EnableMSStandardBlockedWords', value: 'false' },
        { name: 'ClassificationDescriptions', value: '' },
        { name: 'PrefixSuffixNamingRequirement', value: '' },
        { name: 'AllowGuestsToBeGroupOwner', value: 'false' },
        { name: 'AllowGuestsToAccessGroups', value: 'true' },
        { name: 'GuestUsageGuidelinesUrl', value: '' },
        { name: 'GroupCreationAllowedGroupId', value: '' },
        { name: 'AllowToAddGuests', value: 'true' },
        { name: 'EnableGroupCreation', value: 'true' }
      ]
    });
  });

  it('handles error when no template with the specified id found',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation().rejects({
        error: {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "request-id": "fe2491f9-53e7-407c-9a08-b92b2bf6722b",
              "date": "2018-05-11T17:06:22"
            }
          }
        }
      });

      await assert.rejects(command.action(logger, { options: { id: '62375ab9-6b52-47ed-826b-58e47e0e304c' } } as any),
        new CommandError(`Resource '62375ab9-6b52-47ed-826b-58e47e0e304c' does not exist or one of its queried reference-property objects are not present.`));
    }
  );

  it('handles error when group setting already exists', async () => {
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettingTemplates/62375ab9-6b52-47ed-826b-58e47e0e304b`) {
        return {
          "id": "62375ab9-6b52-47ed-826b-58e47e0e304b", "deletedDateTime": null, "displayName": "Group.Unified", "description": "\n        Setting templates define the different settings that can be used for the associated ObjectSettings. This template defines\n        settings that can be used for Unified Groups.\n      ", "values": [{ "name": "CustomBlockedWordsList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of blocked words for Unified Group displayName and mailNickName." }, { "name": "EnableMSStandardBlockedWords", "type": "System.Boolean", "defaultValue": "false", "description": "A flag indicating whether or not to enable the Microsoft Standard list of blocked words for Unified Group displayName and mailNickName." }, { "name": "ClassificationDescriptions", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of structured strings describing the classification values in the ClassificationList. The structure of the string is: Value: Description" }, { "name": "DefaultClassification", "type": "System.String", "defaultValue": "", "description": "The classification value to be used by default for Unified Group creation." }, { "name": "PrefixSuffixNamingRequirement", "type": "System.String", "defaultValue": "", "description": "A structured string describing how a Unified Group displayName and mailNickname should be structured. Please refer to docs to discover how to structure a valid requirement." }, { "name": "AllowGuestsToBeGroupOwner", "type": "System.Boolean", "defaultValue": "false", "description": "Flag indicating if guests are allowed to be owner in any Unified Group." }, { "name": "AllowGuestsToAccessGroups", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed to access any Unified Group resources." }, { "name": "GuestUsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines for guests." }, { "name": "GroupCreationAllowedGroupId", "type": "System.Guid", "defaultValue": "", "description": "Guid of the security group that is always allowed to create Unified Groups." }, { "name": "AllowToAddGuests", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if guests are allowed in any Unified Group." }, { "name": "UsageGuidelinesUrl", "type": "System.String", "defaultValue": "", "description": "A link to the Group Usage Guidelines." }, { "name": "ClassificationList", "type": "System.String", "defaultValue": "", "description": "A comma-delimited list of valid classification values that can be applied to Unified Groups." }, { "name": "EnableGroupCreation", "type": "System.Boolean", "defaultValue": "true", "description": "Flag indicating if group creation feature is on." }]
        };
      }

      throw 'Invalid Request';
    });
    jest.spyOn(request, 'post').mockClear().mockImplementation().rejects({
      error: {
        "error": {
          "code": "Request_BadRequest",
          "message": "A conflicting object with one or more of the specified property values is present in the directory.",
          "innerError": {
            "request-id": "7b7eacbb-3b0e-4758-be20-6410957e42d6",
            "date": "2018-05-11T17:10:34"
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { templateId: '62375ab9-6b52-47ed-826b-58e47e0e304b' } } as any),
      new CommandError(`A conflicting object with one or more of the specified property values is present in the directory.`));
  });

  it('fails validation if the templateId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { templateId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the templateId is a valid GUID', async () => {
    const actual = await command.validate({ options: { templateId: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });
});
