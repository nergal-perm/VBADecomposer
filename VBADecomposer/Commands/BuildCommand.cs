/*
 * Created by SharpDevelop.
 * User: terekhov-ev
 * Date: 31.07.2015
 * Time: 13:31
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace VBADecomposer.Commands {
    public class BuildCommand : BaseCommand {	
		public BuildCommand(string[] commandLine)
			: base(commandLine) {
		}
	
        #region Implemented abstract members

        public override bool run() {
            throw new NotImplementedException("Not yet Implemented");
        }

        public override bool argsAreOk() {
            throw new NotImplementedException("Not yet Implemented");
        }

        #endregion
        
        
    }
}
