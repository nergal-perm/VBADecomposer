/*
 * Created by SharpDevelop.
 * User: terekhov-ev
 * Date: 31.07.2015
 * Time: 9:05
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace VBADecomposer {
	class Program {
		public static void Main(string[] args) {
			Console.WriteLine("Hello World!");
			try {
				ExtractCode();
			} catch (Exception e) {
				Console.WriteLine(e.Message);
				Console.WriteLine(e.StackTrace);
			}
			// TODO: Implement Functionality Here
						
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
		
		public static void ExtractCode() {
			Microsoft.Office.Interop.Excel.Application _xlApp = new Microsoft.Office.Interop.Excel.Application();
			
			// open a workbook with disabled macros
			var tempMacroPolicy = _xlApp.AutomationSecurity;
			_xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
			var wb = _xlApp.Workbooks.Open(@"D:\Temp\Redminereports.xlsm");
			_xlApp.AutomationSecurity = tempMacroPolicy;
			
			// Scan through VBComponents
			var project = wb.VBProject;
			foreach (VBComponent module in project.VBComponents) {
				ExtractVBComponent(module, wb.Path +@"\Source\", "", true);
			}
			wb.Close(false);
			Marshal.ReleaseComObject(wb);
			_xlApp.Quit();
		}
		
		public static void ExtractVBComponent(VBComponent comp, string path, string fName, bool overwrite) {
			// define file name
			var extension = GetFileExtensionFor(comp);
			if (fName.Trim() == String.Empty) {
				fName = comp.Name + extension;
			} else if (fName.IndexOf("."[0]) == 0) {
				fName = fName + extension;
			}
			
			// define folder path			
			if (path.EndsWith(@"\", StringComparison.CurrentCultureIgnoreCase)) {
				fName = path + fName;
			} else {
				fName = path + @"\" + fName;
			}
			
			// is it possible to write to path
			FileInfo file = new FileInfo(fName);
			if (file.Exists) {
				if(overwrite) {
					file.Delete();
				}
			}
			
			comp.Export(fName);
		}
		
		private static string GetFileExtensionFor(VBComponent vbComp) {
			switch (vbComp.Type) {
				case vbext_ComponentType.vbext_ct_ClassModule:
					return ".cls";
				case vbext_ComponentType.vbext_ct_Document:
					return ".cls";
				case vbext_ComponentType.vbext_ct_MSForm:
					return ".frm";
				case vbext_ComponentType.vbext_ct_StdModule:
					return ".bas";
				default:
					return ".bas";
			}
		}
			
	}
}