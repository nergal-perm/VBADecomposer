/*
 * Created by SharpDevelop.
 * User: terekhov-ev
 * Date: 31.07.2015
 * Time: 9:05
 *
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using VBADecomposer.Commands;


namespace VBADecomposer {
	class Program {
		public static void Main(string[] args) {
			BaseCommand cmd = Commands.CommandFactory.getCommand(args);
			try {
				cmd.run();
			} catch (Exception e) {
				Console.WriteLine(e.Message);
				Console.WriteLine(e.StackTrace);
			}

			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}

	}
}
