using System;

namespace VBADecomposer.Commands {
    ///<summary>
    /// Базовый класс для всех команд приложения
    ///</summary>
    public abstract class BaseCommand {
    	protected string[] _commandLine;

    	protected BaseCommand(string[] commandLine) {
    		_commandLine = commandLine;
    	}

    	public abstract bool run();
        public abstract bool argsAreOk();
    }
}
