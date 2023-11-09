import assert from 'assert';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { jestUtil } from '../../../../utils/jestUtil.js';
import commands from '../../commands.js';
import command from './siteclassification-disable.js';

describe(commands.SITECLASSIFICATION_DISABLE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

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
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    jestUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });


  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITECLASSIFICATION_DISABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before disabling siteclassification when confirm option not passed',
    async () => {
      await command.action(logger, { options: {} });
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      assert(promptIssued);
    }
  );

  it('handles Microsoft 365 Tenant siteclassification is not enabled',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return { value: [] };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { debug: true, force: true } } as any),
        new CommandError('Site classification is not enabled.'));
    }
  );

  it('handles Microsoft 365 Tenant siteclassification missing DirectorySettingTemplate',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "d20c475c-6f96-449a-aee8-08146be187d3",
                "displayName": "Group.Unified_not_exist",
                "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
                "values": [
                  {
                    "name": "CustomBlockedWordsList",
                    "value": ""
                  },
                  {
                    "name": "EnableMSStandardBlockedWords",
                    "value": "false"
                  },
                  {
                    "name": "ClassificationDescriptions",
                    "value": ""
                  },
                  {
                    "name": "DefaultClassification",
                    "value": "TopSecret"
                  },
                  {
                    "name": "PrefixSuffixNamingRequirement",
                    "value": ""
                  },
                  {
                    "name": "AllowGuestsToBeGroupOwner",
                    "value": "false"
                  },
                  {
                    "name": "AllowGuestsToAccessGroups",
                    "value": "true"
                  },
                  {
                    "name": "GuestUsageGuidelinesUrl",
                    "value": ""
                  },
                  {
                    "name": "GroupCreationAllowedGroupId",
                    "value": ""
                  },
                  {
                    "name": "AllowToAddGuests",
                    "value": "true"
                  },
                  {
                    "name": "UsageGuidelinesUrl",
                    "value": "https://test"
                  },
                  {
                    "name": "ClassificationList",
                    "value": "TopSecret"
                  },
                  {
                    "name": "EnableGroupCreation",
                    "value": "true"
                  }
                ]
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { debug: true, force: true } } as any),
        new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
    }
  );

  it('handles Microsoft 365 Tenant siteclassification missing UnifiedGroupSetting ID',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id_doesnotexists": "d20c475c-6f96-449a-aee8-08146be187d3",
                "displayName": "Group.Unified",
                "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
                "values": [
                  {
                    "name": "CustomBlockedWordsList",
                    "value": ""
                  },
                  {
                    "name": "EnableMSStandardBlockedWords",
                    "value": "false"
                  },
                  {
                    "name": "ClassificationDescriptions",
                    "value": ""
                  },
                  {
                    "name": "DefaultClassification",
                    "value": "TopSecret"
                  },
                  {
                    "name": "PrefixSuffixNamingRequirement",
                    "value": ""
                  },
                  {
                    "name": "AllowGuestsToBeGroupOwner",
                    "value": "false"
                  },
                  {
                    "name": "AllowGuestsToAccessGroups",
                    "value": "true"
                  },
                  {
                    "name": "GuestUsageGuidelinesUrl",
                    "value": ""
                  },
                  {
                    "name": "GroupCreationAllowedGroupId",
                    "value": ""
                  },
                  {
                    "name": "AllowToAddGuests",
                    "value": "true"
                  },
                  {
                    "name": "UsageGuidelinesUrl",
                    "value": "https://test"
                  },
                  {
                    "name": "ClassificationList",
                    "value": "TopSecret"
                  },
                  {
                    "name": "EnableGroupCreation",
                    "value": "true"
                  }
                ]
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { debug: true, force: true } } as any),
        new CommandError("Missing UnifiedGroupSettting id"));
    }
  );

  it('handles Microsoft 365 Tenant siteclassification empty UnifiedGroupSetting ID',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "",
                "displayName": "Group.Unified",
                "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
                "values": [
                  {
                    "name": "CustomBlockedWordsList",
                    "value": ""
                  },
                  {
                    "name": "EnableMSStandardBlockedWords",
                    "value": "false"
                  },
                  {
                    "name": "ClassificationDescriptions",
                    "value": ""
                  },
                  {
                    "name": "DefaultClassification",
                    "value": "TopSecret"
                  },
                  {
                    "name": "PrefixSuffixNamingRequirement",
                    "value": ""
                  },
                  {
                    "name": "AllowGuestsToBeGroupOwner",
                    "value": "false"
                  },
                  {
                    "name": "AllowGuestsToAccessGroups",
                    "value": "true"
                  },
                  {
                    "name": "GuestUsageGuidelinesUrl",
                    "value": ""
                  },
                  {
                    "name": "GroupCreationAllowedGroupId",
                    "value": ""
                  },
                  {
                    "name": "AllowToAddGuests",
                    "value": "true"
                  },
                  {
                    "name": "UsageGuidelinesUrl",
                    "value": "https://test"
                  },
                  {
                    "name": "ClassificationList",
                    "value": "TopSecret"
                  },
                  {
                    "name": "EnableGroupCreation",
                    "value": "true"
                  }
                ]
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      await assert.rejects(command.action(logger, { options: { debug: true, force: true } } as any),
        new CommandError("Missing UnifiedGroupSettting id"));
    }
  );

  it('handles disabling site classification without prompting', async () => {
    let deleteRequestIssued = false;
    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return {
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/d20c475c-6f96-449a-aee8-08146be187d3`) {
        deleteRequestIssued = true;
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true } } as any);
    assert(deleteRequestIssued);
  });

  it('handles disabling site classification without prompting (debug)',
    async () => {
      let deleteRequestIssued = false;
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "d20c475c-6f96-449a-aee8-08146be187d3",
                "displayName": "Group.Unified",
                "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
                "values": [
                  {
                    "name": "CustomBlockedWordsList",
                    "value": ""
                  },
                  {
                    "name": "EnableMSStandardBlockedWords",
                    "value": "false"
                  },
                  {
                    "name": "ClassificationDescriptions",
                    "value": ""
                  },
                  {
                    "name": "DefaultClassification",
                    "value": "TopSecret"
                  },
                  {
                    "name": "PrefixSuffixNamingRequirement",
                    "value": ""
                  },
                  {
                    "name": "AllowGuestsToBeGroupOwner",
                    "value": "false"
                  },
                  {
                    "name": "AllowGuestsToAccessGroups",
                    "value": "true"
                  },
                  {
                    "name": "GuestUsageGuidelinesUrl",
                    "value": ""
                  },
                  {
                    "name": "GroupCreationAllowedGroupId",
                    "value": ""
                  },
                  {
                    "name": "AllowToAddGuests",
                    "value": "true"
                  },
                  {
                    "name": "UsageGuidelinesUrl",
                    "value": "https://test"
                  },
                  {
                    "name": "ClassificationList",
                    "value": "TopSecret"
                  },
                  {
                    "name": "EnableGroupCreation",
                    "value": "true"
                  }
                ]
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/d20c475c-6f96-449a-aee8-08146be187d3`) {
          deleteRequestIssued = true;
          return { value: [] };
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, force: true } } as any);
      assert(deleteRequestIssued);
    }
  );

  it('aborts removing the group when prompt not confirmed', async () => {
    const postSpy = jest.spyOn(request, 'delete').mockClear();
    jestUtil.restore(Cli.prompt);
    jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: false });

    await command.action(logger, { options: {} });
    assert(postSpy.notCalled);
  });

  it('handles disabling site classification when prompt confirmed',
    async () => {
      let deleteRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "d20c475c-6f96-449a-aee8-08146be187d3",
                "displayName": "Group.Unified",
                "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
                "values": [
                  {
                    "name": "CustomBlockedWordsList",
                    "value": ""
                  },
                  {
                    "name": "EnableMSStandardBlockedWords",
                    "value": "false"
                  },
                  {
                    "name": "ClassificationDescriptions",
                    "value": ""
                  },
                  {
                    "name": "DefaultClassification",
                    "value": "TopSecret"
                  },
                  {
                    "name": "PrefixSuffixNamingRequirement",
                    "value": ""
                  },
                  {
                    "name": "AllowGuestsToBeGroupOwner",
                    "value": "false"
                  },
                  {
                    "name": "AllowGuestsToAccessGroups",
                    "value": "true"
                  },
                  {
                    "name": "GuestUsageGuidelinesUrl",
                    "value": "https://test"
                  },
                  {
                    "name": "GroupCreationAllowedGroupId",
                    "value": ""
                  },
                  {
                    "name": "AllowToAddGuests",
                    "value": "true"
                  },
                  {
                    "name": "UsageGuidelinesUrl",
                    "value": "https://test"
                  },
                  {
                    "name": "ClassificationList",
                    "value": "TopSecret"
                  },
                  {
                    "name": "EnableGroupCreation",
                    "value": "true"
                  }
                ]
              }
            ]
          };
        }

        throw 'Invalid request';
      });

      jest.spyOn(request, 'delete').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/d20c475c-6f96-449a-aee8-08146be187d3`) {
          deleteRequestIssued = true;
          return { value: [] };
        }

        throw 'Invalid request';
      });

      jestUtil.restore(Cli.prompt);
      jest.spyOn(Cli, 'prompt').mockClear().mockImplementation().resolves({ continue: true });

      await command.action(logger, { options: {} });
      assert(deleteRequestIssued);
    }
  );
});
