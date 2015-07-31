﻿/*
 * Created by SharpDevelop.
 * User: terekhov-ev
 * Date: 31.07.2015
 * Time: 13:27
 *
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using NUnit.Framework;
using VBADecomposer.Commands;

namespace VBADecomposer.Test {
	[TestFixture]
	public class CommandDispatcherTest {
		BaseCommand _command;

		[Test]
		public void shouldGetCorrectDecomposeCommand() {
			// -file parameter (with value = path to workbook) is required
			// If it's not present or its value isn't set then return HelpCommand to show usage reference
			_command = CommandFactory.getCommand(new string[] { "decompose" });
			Assert.AreEqual(typeof(HelpCommand), _command.GetType());

			_command = CommandFactory.getCommand(new string[] {"decompose", "something", "-file"});
			Assert.AreEqual(typeof(HelpCommand), _command.GetType());

			_command = CommandFactory.getCommand(new string[] { "decompose", "-file", @"D:\temp\RedmineReports.xlsm" });
			Assert.AreEqual(typeof(DecomposeCommand), _command.GetType());
			Assert.IsTrue(((DecomposeCommand)_command).CheckPrerequisites());

			_command = CommandFactory.getCommand(new string[] { "decompose", "-key", "value",
				"-file", @"D:\temp\RedmineReports.xlsm" });
			Assert.AreEqual(typeof(DecomposeCommand), _command.GetType());
			Assert.IsTrue(((DecomposeCommand)_command).CheckPrerequisites());
			
			_command = CommandFactory.getCommand(new string[] { "decompose", "-key", "value",
				"-file", @"D:\temp1\RedmineReports.xlsm" });
			Assert.AreEqual(typeof(DecomposeCommand), _command.GetType());
			Assert.IsFalse(((DecomposeCommand)_command).CheckPrerequisites());			
		}
	}
}