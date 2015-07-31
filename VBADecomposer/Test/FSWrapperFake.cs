/*
 * Created by SharpDevelop.
 * User: terekhov-ev
 * Date: 31.07.2015
 * Time: 11:48
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using VBADecomposer.Main;

namespace VBADecomposer.Test {
	/// <summary>
	/// Description of FSWrapperFake.
	/// </summary>
	public class FSWrapperFake : FSWrapper {
		public FSWrapperFake() {
		}
		
		public string WorkbookPath { get; set; }
		
		public override bool WorkbookExists() {
			switch (WorkbookPath) {
				case @"C:\existing_workbook.xls":
					return true;
				case @"C:\non_existing_workbook.xls":
					return false;
				default:
					return false;
			}
		}
	}
}
