using System;

namespace VBADecomposer.Commands {
	/// <summary>
	/// Description of CommandFactory.
	/// </summary>
	public static class CommandFactory {
		public static BaseCommand getCommand(string[] commandLine) {
			BaseCommand command = null;

			if (commandLine.Length != 0) {
				switch (commandLine[0].ToUpper()) {
					case "DECOMPOSE":
						command = new DecomposeCommand(commandLine);
						break;
					case "BUILD":
						command = new BuildCommand(commandLine);
						break;
					default:
						command = new HelpCommand(commandLine);
						break;
				}
			}

			if (command == null || !command.argsAreOk()) {
				command = new HelpCommand(commandLine);
			}

			return command;
		}
	}
}
