import { Cli, CommandOutput } from "./cli";
import * as path from 'path';
import auth, { AuthType } from './Auth';

export function executeCommand(commandName: string, options: any, listener?: {
  stdout: (message: any) => void,
  stderr: (message: any) => void,
}): Promise<CommandOutput> {
  const cli = Cli.getInstance();
  cli.commandsFolder = path.join(__dirname, 'm365');
  cli.commands = [];
  cli.loadCommandFromArgs(commandName.split(' '));
  if (cli.commands.length !== 1) {
    return Promise.reject(`Command not found: ${commandName}`);
  }

  return Cli.executeCommandWithOutput(cli.commands[0].command, { options: options ?? {} }, listener);
}

export function loginWithCookie(cookie: string): void {
  auth.service.logout();
  auth.service.authType = AuthType.Cookie;
  auth.service.connected = true;
  auth.service.cookie = cookie;
}