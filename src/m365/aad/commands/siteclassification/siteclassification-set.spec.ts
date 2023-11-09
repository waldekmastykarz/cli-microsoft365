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
import command from './siteclassification-set.js';

describe(commands.SITECLASSIFICATION_SET, () => {
  let log: string[];
  let logger: Logger;
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
  });

  afterEach(() => {
    jestUtil.restore([
      request.patch,
      request.get
    ]);
  });

  afterAll(() => {
    jest.restoreAllMocks();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITECLASSIFICATION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if none of the options are specified', async () => {
    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if at least one option is specified', async () => {
    const actual = await command.validate({
      options: {
        classifications: "Confidential"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all options are passed', async () => {
    const actual = await command.validate({
      options: {
        classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "https://aka.ms/pnp", guestUsageGuidelinesUrl: "https://aka.ms/pnp"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles Microsoft 365 Tenant siteclassification has not been enabled',
    async () => {
      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: []
          };
        }

        throw 'Invalid Request';
      });

      await assert.rejects(command.action(logger, { options: { debug: true, classifications: "HBI, LBI, Top Secret", defaultClassification: "HBI", usageGuidelinesUrl: "http://aka.ms/sppnp" } } as any),
        new CommandError("There is no previous defined site classification which can updated."));
    }
  );

  it('updates Microsoft 365 Tenant usage guidelines url and guest usage guidelines url (debug)',
    async () => {
      let updateRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                    "value": "middle"
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
                    "value": ""
                  },
                  {
                    "name": "ClassificationList",
                    "value": "high,middle,low"
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

        throw 'Invalid Request';
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
          JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
          updateRequestIssued = true;

          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { debug: true, usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } } as any);
      assert(updateRequestIssued);
    }
  );

  it('updates Microsoft 365 Tenant usage guidelines url and guest usage guidelines url',
    async () => {
      let updateRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                    "value": "middle"
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
                    "value": ""
                  },
                  {
                    "name": "ClassificationList",
                    "value": "high,middle,low"
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

        throw 'Invalid Request';
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
          JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
          updateRequestIssued = true;

          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } } as any);
      assert(updateRequestIssued);
    }
  );

  it('updates Microsoft 365 Tenant usage guidelines url', async () => {
    let updateRequestIssued = false;

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return {
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": "high,middle,low"
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

      throw 'Invalid Request';
    });

    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { usageGuidelinesUrl: "http://aka.ms/pnp" } } as any);
    assert(updateRequestIssued);
  });

  it('updates Microsoft 365 Tenant guest usage guidelines url', async () => {
    let updateRequestIssued = false;

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return {
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": "high,middle,low"
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

      throw 'Invalid Request';
    });

    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { guestUsageGuidelinesUrl: "http://aka.ms/pnp" } } as any);
    assert(updateRequestIssued);
  });

  it('updates Microsoft 365 Tenant siteclassification', async () => {
    let updateRequestIssued = false;

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return {
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": "high,middle,low"
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

      throw 'Invalid Request';
    });

    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"middle"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"top secret,high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;

        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { classifications: "top secret,high,middle,low" } } as any);
    assert(updateRequestIssued);
  });

  it('updates Microsoft 365 Tenant default classification', async () => {
    let updateRequestIssued = false;

    jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return {
          value: [
            {
              "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                  "value": "middle"
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": "high,middle,low"
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

      throw 'Invalid Request';
    });

    jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
        JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"low"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
        updateRequestIssued = true;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { defaultClassification: "low" } } as any);
    assert(updateRequestIssued);
  });

  it('updates Microsoft 365 Tenant siteclassification and default classification',
    async () => {
      let updateRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                    "value": "middle"
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
                    "value": ""
                  },
                  {
                    "name": "ClassificationList",
                    "value": "high,middle,low"
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

        throw 'Invalid Request';
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
          JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"high"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":"area 51,high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
          updateRequestIssued = true;

          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { classifications: "area 51,high,middle,low", defaultClassification: "high" } } as any);
      assert(updateRequestIssued);
    }
  );

  it('updates Microsoft 365 Tenant siteclassification, default classification, usage guidelines url and guest usage guidelines url',
    async () => {
      let updateRequestIssued = false;

      jest.spyOn(request, 'get').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
          return {
            value: [
              {
                "id": "a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b",
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
                    "value": "middle"
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
                    "value": ""
                  },
                  {
                    "name": "ClassificationList",
                    "value": "high,middle,low"
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

        throw 'Invalid Request';
      });

      jest.spyOn(request, 'patch').mockClear().mockImplementation(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings/a557c1d2-ef9d-4ac5-ad45-7f8b22d9250b` &&
          JSON.stringify(opts.data) === `{"values":[{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":"high"},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":"http://aka.ms/pnp"},{"name":"ClassificationList","value":"area 51,high,middle,low"},{"name":"EnableGroupCreation","value":"true"}]}`) {
          updateRequestIssued = true;

          return;
        }

        throw 'Invalid request';
      });

      await command.action(logger, { options: { classifications: "area 51,high,middle,low", defaultClassification: "high", usageGuidelinesUrl: "http://aka.ms/pnp", guestUsageGuidelinesUrl: "http://aka.ms/pnp" } } as any);
      assert(updateRequestIssued);
    }
  );
});
