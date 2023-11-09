import assert from 'assert';
import fs from 'fs';
import inquirer from 'inquirer';
import { chili } from './chili.js';
import request from '../request.js';
import { jestUtil } from '../utils/jestUtil.js';

describe('chili', () => {
  let consoleLogSpy: jest.Mock;
  let consoleErrorSpy: jest.Mock;

  beforeAll(() => {
    consoleLogSpy = jest.spyOn(console, 'log').mockClear().mockImplementation(() => { });
    consoleErrorSpy = jest.spyOn(console, 'error').mockClear().mockImplementation(() => { });
  });

  afterEach(() => {
    consoleLogSpy.mockReset();
    consoleErrorSpy.mockReset();

    jestUtil.restore([
      inquirer.prompt,
      request.post,
      fs.existsSync
    ]);
  });

  afterAll(() => {
    jestUtil.restore([
      // eslint-disable-next-line no-console
      console.log,
      // eslint-disable-next-line no-console
      console.error
    ]);
  });

  it('starts a conversation using a prompt from args when specified',
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
        switch (options.url) {
          case 'https://api.mendable.ai/v0/newConversation':
            return Promise.resolve({
              // eslint-disable-next-line camelcase
              conversation_id: 1
            });
          case 'https://api.mendable.ai/v0/mendableChat':
            if (options.data.question === 'Hello') {
              return Promise.resolve({
                answer: {
                  text: 'Hello back'
                },
                sources: []
              });
            }
            break;
        }
        return Promise.reject('Invalid request');
      });
      jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
      assert.doesNotReject(chili.startConversation(['Hello']));
    }
  );

  it('starts a conversation when a prompt specified as a single arg', () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
    assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it('starts a conversation when a prompt specified as multiple args (no quotes)',
    () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
        switch (options.url) {
          case 'https://api.mendable.ai/v0/newConversation':
            return Promise.resolve({
              // eslint-disable-next-line camelcase
              conversation_id: 1
            });
          case 'https://api.mendable.ai/v0/mendableChat':
            if (options.data.question === 'Hello world') {
              return Promise.resolve({
                answer: {
                  text: 'Hello back'
                },
                sources: []
              });
            }
            break;
        }
        return Promise.reject('Invalid request');
      });
      jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
      assert.doesNotReject(chili.startConversation(['Hello', 'world']));
    }
  );

  it('starts a conversation asking for a prompt when no prompt specified via args',
    () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
        switch (options.url) {
          case 'https://api.mendable.ai/v0/newConversation':
            return Promise.resolve({
              // eslint-disable-next-line camelcase
              conversation_id: 1
            });
          case 'https://api.mendable.ai/v0/mendableChat':
            if (options.data.question === 'Hello world') {
              return Promise.resolve({
                answer: {
                  text: 'Hello back'
                },
                sources: []
              });
            }
            break;
        }
        return Promise.reject('Invalid request');
      });
      jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
        if (questions[0].name === 'prompt') {
          return Promise.resolve({
            prompt: 'Hello world'
          });
        }
        return Promise.resolve({});
      });
      assert.doesNotReject(chili.startConversation([]));
    }
  );

  it('uses the prompt to search in CLI docs using Mendable', () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
    assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('displays the retrieved response to the user', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
    await chili.startConversation(['Hello']);
    assert(consoleLogSpy.calledWith('Hello back'));
  });

  it('in response formats MD in terminal-friendly way', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello **back**'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
    await chili.startConversation(['Hello']);
    assert(consoleLogSpy.calledWith('Hello back'));
  });

  it('in response, shows sources', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: [
                {
                  link: 'https://example.com/source-1'
                }
              ]
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
    await chili.startConversation(['Hello']);
    assert(consoleLogSpy.calledWith('⬥ https://example.com/source-1'));
  });

  it('prompts for rating the response when rating is enabled', async () => {
    let promptedForRating = false;
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      if (questions[0].name === 'prompt') {
        return Promise.resolve({
          prompt: 'Hello world'
        });
      }
      if (questions[0].name === 'rating') {
        promptedForRating = true;
      }
      return Promise.resolve({});
    });
    await chili.startConversation(['Hello world']);
    assert.strictEqual(promptedForRating, true);
  });

  it(`doesn't prompt for rating the response when rating is disabled`,
    async () => {
      let promptedForRating = false;
      jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
        switch (options.url) {
          case 'https://api.mendable.ai/v0/newConversation':
            return Promise.resolve({
              // eslint-disable-next-line camelcase
              conversation_id: 1
            });
          case 'https://api.mendable.ai/v0/mendableChat':
            if (options.data.question === 'Hello world') {
              return Promise.resolve({
                answer: {
                  text: 'Hello back'
                },
                // eslint-disable-next-line camelcase
                message_id: 1,
                sources: []
              });
            }
            break;
        }
        return Promise.reject('Invalid request');
      });
      jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
        if (questions[0].name === 'prompt') {
          return Promise.resolve({
            prompt: 'Hello world'
          });
        }
        if (questions[0].name === 'rating') {
          promptedForRating = true;
        }
        return Promise.resolve({});
      });
      await chili.startConversation(['Hello world', '--no-rating']);
      assert.strictEqual(promptedForRating, false);
    }
  );

  it('sends positive rating to Mendable', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          if (options.data.rating === 1) {
            return Promise.resolve({});
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      if (questions[0].name === 'prompt') {
        return Promise.resolve({
          prompt: 'Hello world'
        });
      }
      if (questions[0].name === 'rating') {
        return Promise.resolve({
          rating: 1
        });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it('sends negative rating to Mendable', () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          if (options.data.rating === -1) {
            return Promise.resolve({});
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      if (questions[0].name === 'prompt') {
        return Promise.resolve({
          prompt: 'Hello world'
        });
      }
      if (questions[0].name === 'rating') {
        return Promise.resolve({
          rating: -1
        });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it(`doesn't send rating to Mendable when user chose to skip`, async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      if (questions[0].name === 'prompt') {
        return Promise.resolve({
          prompt: 'Hello world'
        });
      }
      if (questions[0].name === 'rating') {
        return Promise.resolve({
          rating: 0
        });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it(`doesn't fail when rating the response failed`, async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      if (questions[0].name === 'prompt') {
        return Promise.resolve({
          prompt: 'Hello world'
        });
      }
      if (questions[0].name === 'rating') {
        return Promise.resolve({
          rating: 1
        });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it(`when rating the response failed, shows error in debug mode`,
    async () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
        switch (options.url) {
          case 'https://api.mendable.ai/v0/newConversation':
            return Promise.resolve({
              // eslint-disable-next-line camelcase
              conversation_id: 1
            });
          case 'https://api.mendable.ai/v0/mendableChat':
            if (options.data.question === 'Hello world') {
              return Promise.resolve({
                answer: {
                  text: 'Hello back'
                },
                // eslint-disable-next-line camelcase
                message_id: 1,
                sources: []
              });
            }
            break;
          case 'https://api.mendable.ai/v0/rateMessage':
            return Promise.reject('An error has occurred');
        }
        return Promise.reject('Invalid request');
      });
      jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
        if (questions[0].name === 'prompt') {
          return Promise.resolve({
            prompt: 'Hello world'
          });
        }
        if (questions[0].name === 'rating') {
          return Promise.resolve({
            rating: 1
          });
        }
        return Promise.resolve({});
      });
      await chili.startConversation(['Hello world', '--debug']);
      assert(consoleErrorSpy.calledWith('An error has occurred while rating the response: An error has occurred'));
    }
  );

  it('allows asking a follow-up question after a response', async () => {
    let i = 0;
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          if (options.data.question === 'Follow up') {
            return Promise.resolve({
              answer: {
                text: 'Hello again'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      switch (questions[0].name) {
        case 'chat':
          if (i++ === 0) {
            return Promise.resolve({
              chat: 'ask'
            });
          }
          else {
            return Promise.resolve({
              chat: 'end'
            });
          }
        case 'prompt':
          return Promise.resolve({
            prompt: 'Follow up'
          });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('for a follow-up question, includes the history', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          if (options.data.question === 'Follow up' &&
            options.data.history[0].prompt === 'Hello' &&
            options.data.history[0].answer === 'Hello back') {
            return Promise.resolve({
              answer: {
                text: 'Hello again'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      switch (questions[0].name) {
        case 'chat':
          return Promise.resolve({
            chat: 'ask'
          });
        case prompt:
          return Promise.resolve({
            prompt: 'Follow up'
          });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('allows ending conversation after a response', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      if (questions[0].name === 'chat') {
        return Promise.resolve({
          chat: 'end'
        });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('allows starting a new conversation after a response', async () => {
    let i = 0;
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: ++i
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          if (options.data.question === 'Hello 2' &&
            options.data.conversation_id === 2) {
            return Promise.resolve({
              answer: {
                text: 'Hello there'
              },
              sources: []
            });
          }
          break;
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation((questions: any): any => {
      switch (questions[0].name) {
        case 'chat':
          return Promise.resolve({
            chat: 'new'
          });
        case 'prompt':
          return Promise.resolve({
            prompt: 'Hello 2'
          });
      }
      return Promise.resolve({});
    });
    assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('throws exception when getting conversation ID failed', async () => {
    jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });
    jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
    assert.rejects(chili.startConversation(['Hello']), 'An error has occurred');
  });

  it('throw exception when calling Mendable API to search in CLI docs failed',
    () => {
      jest.spyOn(request, 'post').mockClear().mockImplementation(options => {
        switch (options.url) {
          case 'https://api.mendable.ai/v0/newConversation':
            return Promise.resolve({
              // eslint-disable-next-line camelcase
              conversation_id: 1
            });
          case 'https://api.mendable.ai/v0/mendableChat':
            return Promise.reject('An error has occurred');
        }
        return Promise.reject('Invalid request');
      });
      jest.spyOn(inquirer, 'prompt').mockClear().mockImplementation().resolves({});
      assert.rejects(chili.startConversation(['Hello']), 'An error has occurred');
    }
  );

  it('shows help when requested using --help', async () => {
    await chili.startConversation(['--help']);
    assert(consoleLogSpy.mock.calls.some(call => call.mock.calls[0].includes('CLI ASSISTANT (CHILI)')));
  });

  it('shows help when requested using -h', async () => {
    await chili.startConversation(['-h']);
    assert(consoleLogSpy.mock.calls.some(call => call.mock.calls[0].includes('CLI ASSISTANT (CHILI)')));
  });

  it(`when requested help, doesn't start conversation`, async () => {
    const requestSpy = jest.spyOn(request, 'post').mockClear();
    await chili.startConversation(['--help']);
    assert(requestSpy.notCalled);
  });

  it('shows error message when help file not found', async () => {
    jest.spyOn(fs, 'existsSync').mockClear().mockImplementation(() => false);
    await chili.startConversation(['--help']);
    assert(consoleErrorSpy.calledWith('Help file not found'));
  });
});