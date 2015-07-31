using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace VBADecomposer.Commands {
    public sealed class DecomposeCommand : BaseCommand {
		private string _workbookPath;


		public DecomposeCommand(string[] commandLine)
			: base(commandLine) {
		}

        #region Implemented abstract members

        public override bool run() {
        	if (!CheckPrerequisites()) {
        		return false;
        	}
        	ExtractCode();
            return true;
        }

        public override bool argsAreOk() {
            bool fileParam = false; //-file parameter is required, so we keep a flag to show if it's present
        	int argsCount = _commandLine.Length;

        	for (int i = 1; i < argsCount; i++) {
        		switch (_commandLine[i].ToUpper()) {
        			case "-FILE":
        				if (i+1 < argsCount) {
                            _workbookPath = _commandLine[i+1];
                            fileParam = true;
                        } else {
                            return false;
                        }
                        break;
                    default:
                        break;
        		}
        	}

        	return fileParam;
        }

        #endregion

		public void ExtractCode() {
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

			// free the workbook and close Excel application
			wb.Close(false);
			Marshal.ReleaseComObject(wb);
			_xlApp.Quit();
		}

		public void ExtractVBComponent(VBComponent comp, string path, string fName, bool overwrite) {
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
        
        public bool CheckPrerequisites() {
        	bool result = true;
        	
        	// 
        	if (!File.Exists(_workbookPath)) {
        		Console.WriteLine("Рабочая книга Excel (" + _workbookPath + ") не найдена!");
        		result = false;
        	}
        	
        	return result;
        }
    }
}
