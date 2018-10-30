import appInsights from './appInsights';
import GlobalOptions from './GlobalOptions';

const vorpal: Vorpal = require('./vorpal-init');

export interface CommandOption {
  option: string;
  description: string;
  autocomplete?: string[]
}

export interface CommandAction {
  (this: CommandInstance, args: any, cb: (err?: any) => void): void
}

export interface CommandValidate {
  (args: any): boolean | string
}

export interface CommandHelp {
  (args: any, cbOrLog: (msg?: string) => void): void
}

export interface CommandCancel {
  (): void
}

export interface CommandTypes {
  string?: string[];
  boolean?: string[];
}

export class CommandError {
  constructor(public message: string, public code?: number) {
  }
}

export interface ODataError {
  "odata.error": {
    code: string;
    message: {
      lang: string;
      value: string;
    }
  }
}

interface CommandArgs {
  options: GlobalOptions;
  stdin?: any;
}

export default abstract class Command {
  private _debug: boolean = false;
  private _verbose: boolean = false;
  private _stdin: any;
  private stdinParsed: boolean = false;

  protected get debug(): boolean {
    return this._debug;
  }

  protected get verbose(): boolean {
    return this._verbose;
  }

  protected get stdin(): any {
    return this._stdin;
  }

  public abstract get name(): string;
  public abstract get description(): string;

  public abstract commandAction(cmd: CommandInstance, args: any, cb: () => void): void;
  public abstract commandHelp(args: any, log: (message: string) => void): void;

  protected showDeprecationWarning(cmd: CommandInstance, deprecated: string, recommended: string): void {
    if (cmd.commandWrapper.command.indexOf(deprecated) === 0) {
      cmd.log(vorpal.chalk.yellow(`Command '${deprecated}' is deprecated. Please use '${recommended}' instead`));
    }
  }

  protected getUsedCommandName(cmd: CommandInstance): string {
    const commandName: string = this.getCommandName();
    if (cmd.commandWrapper.command.indexOf(commandName) === 0) {
      return commandName;
    }

    // since the command was called by something else than its name
    // it must have aliases
    const aliases: string[] = this.alias() as string[];

    for (let i: number = 0; i < aliases.length; i++) {
      if (cmd.commandWrapper.command.indexOf(aliases[i]) === 0) {
        return aliases[i];
      }
    }

    // shouldn't happen because the command is called either by its name or alias
    return '';
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: () => void) {
      cmd.initAction(args);
      cmd.commandAction(this, args, cb);
    }
  }

  public getTelemetryProperties(args: any): any {
    return {
      debug: this.debug.toString(),
      verbose: this.verbose.toString()
    };
  }

  public alias(): string[] | undefined {
    return;
  }

  public autocomplete(): string[] | undefined {
    return;
  }

  public allowUnknownOptions(): boolean | undefined {
    return;
  }

  public options(): CommandOption[] {
    return [
      {
        option: '-o, --output [output]',
        description: 'Output type. json|text. Default text',
        autocomplete: ['json', 'text']
      },
      {
        option: '--verbose',
        description: 'Runs command with verbose logging'
      },
      {
        option: '--debug',
        description: 'Runs command with debug logging'
      }
    ];
  }

  public help(): CommandHelp {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cbOrLog: () => void) {
      const ranFromHelpCommand: boolean =
        typeof vorpal._command !== 'undefined' &&
        typeof vorpal._command.command !== 'undefined' &&
        vorpal._command.command.indexOf('help ') === 0;

      const log = ranFromHelpCommand ? cbOrLog : this.log.bind(this);

      cmd.commandHelp(args, log);

      if (!ranFromHelpCommand) {
        cbOrLog();
      }
    }
  }

  public validate(): CommandValidate | undefined {
    return;
  }

  public cancel(): CommandCancel | undefined {
    return;
  }

  public types(): CommandTypes | undefined {
    return;
  }

  public init(vorpal: Vorpal): void {
    const cmd: VorpalCommand = vorpal
      .command(this.name, this.description, this.autocomplete())
      .action(this.action());
    const options: CommandOption[] = this.options();
    options.forEach((o: CommandOption): void => {
      cmd.option(o.option, o.description, o.autocomplete);
    });
    const alias: string[] | undefined = this.alias();
    if (alias) {
      cmd.alias(alias);
    }
    const validate: CommandValidate | undefined = this.validate();
    if (validate) {
      cmd.validate(validate);
    }
    const cancel: CommandCancel | undefined = this.cancel();
    if (cancel) {
      cmd.cancel(cancel);
    }
    const allowUnknownOptions: boolean | undefined = this.allowUnknownOptions();
    if (allowUnknownOptions) {
      cmd.allowUnknownOptions();
    }
    cmd.help(this.help());
    const types: CommandTypes | undefined = this.types();
    if (types) {
      cmd.types(types);
    }
  }

  public getCommandName(): string {
    let commandName: string = this.name;
    let pos: number = commandName.indexOf('<');
    let pos1: number = commandName.indexOf('[');
    if (pos > -1 || pos1 > -1) {
      if (pos1 > -1) {
        pos = pos1;
      }

      commandName = commandName.substr(0, pos).trim();
    }

    return commandName;
  }

  protected handleRejectedODataPromise(rawResponse: any, cmd: CommandInstance, callback: (err?: any) => void): void {
    const res: any = JSON.parse(JSON.stringify(rawResponse));
    if (res.error) {
      try {
        const err: ODataError = JSON.parse(res.error);
        callback(new CommandError(err['odata.error'].message.value));
      }
      catch {
        callback(new CommandError(res.error));
      }
    }
    else {
      if (rawResponse instanceof Error) {
        callback(new CommandError(rawResponse.message));
      }
      else {
        callback(new CommandError(rawResponse));
      }
    }
  }

  protected handleRejectedODataJsonPromise(response: any, cmd: CommandInstance, callback: (err?: any) => void): void {
    if (response.error &&
      response.error['odata.error'] &&
      response.error['odata.error'].message) {
      callback(new CommandError(response.error['odata.error'].message.value));
    }
    else {
      if (response.error) {
        if (response.error.error &&
          response.error.error.message) {
          callback(new CommandError(response.error.error.message));
        }
        else {
          if (response.error.message) {
            callback(new CommandError(response.error.message));
          }
          else {
            if (response.error.error_description) {
              callback(new CommandError(response.error.error_description));
            }
            else {
              try {
                const error: any = JSON.parse(response.error);
                if (error &&
                  error.error &&
                  error.error.message) {
                  callback(new CommandError(error.error.message));
                }
                else {
                  callback(new CommandError(response.error));
                }
              }
              catch {
                callback(new CommandError(response.error));
              }
            }
          }
        }
      }
      else {
        if (response instanceof Error) {
          callback(new CommandError(response.message));
        }
        else {
          callback(new CommandError(response));
        }
      }
    }
  }

  protected handleError(rawResponse: any, cmd: CommandInstance, callback: (err?: any) => void): void {
    if (rawResponse instanceof Error) {
      callback(new CommandError(rawResponse.message));
    }
    else {
      callback(new CommandError(rawResponse));
    }
  }

  protected handleRejectedPromise(rawResponse: any, cmd: CommandInstance, callback: (err?: any) => void): void {
    this.handleError(rawResponse, cmd, callback);
  }

  protected parseStdin(args: CommandArgs): void {
    // stdin could be parsed multiple times (eg. during command validation or
    // execution). Ensure it's done only once
    if (this.stdinParsed) {
      return;
    }
    this.stdinParsed = true;

    if (!args.stdin) {
      return;
    }

    this._stdin = args.stdin;
    // When running the CLI in the immersive mode, value from the pipe is
    // passed as an array with one element
    if (Array.isArray(this._stdin) && this._stdin.length > 0) {
      this._stdin = this._stdin[0];
    }

    if (typeof this._stdin === 'string') {
      // if the value is a JSON string, deserialize it to object
      try {
        this._stdin = JSON.parse(this._stdin);
      }
      catch { }
    }
  }

  protected initAction(args: CommandArgs): void {
    this._debug = args.options.debug || process.env.OFFICE365CLI_DEBUG === '1';
    this._verbose = this._debug || args.options.verbose || process.env.OFFICE365CLI_VERBOSE === '1';
    this.parseStdin(args);

    appInsights.trackEvent({
      name: this.getCommandName(),
      properties: this.getTelemetryProperties(args)
    });
    appInsights.flush();
  }
}