/*
 * Created by SharpDevelop.
 * User: terekhov-ev
 * Date: 31.07.2015
 * Time: 13:31
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace VBADecomposer.Commands {
    public class BuildCommand : BaseCommand {	
		public BuildCommand(string[] commandLine)
			: base(commandLine) {
		}
	
        #region Implemented abstract members

        public override bool run() {
        	ImportCode();
        	return true;
        }

        public override bool argsAreOk() {
        	return true;
        }

        #endregion
        
        public void ImportCode() {
        	Excel.Application _xlApp = new Excel.Application();
        	Excel.Workbook wb = _xlApp.Workbooks.Add();
        	
        	foreach (FileInfo file in new DirectoryInfo(@"D:\Temp\Source").GetFiles()) {
        		ImportComponent(file, wb);
        	}
        	
        	wb.SaveAs(@"D:\Temp\Built.xlsm", Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
			wb.Close(false);
			Marshal.ReleaseComObject(wb);
			_xlApp.Quit();        	
        }
        
        
        private void ImportComponent(FileInfo f, Excel.Workbook wb) {
        	VBA.VBProject project = wb.VBProject;
        	VBA.VBComponent component = null;
        	try {
        		component = project.VBComponents.Item(f.Name);
        	} catch (Exception e) {
        		// do nothing
        	}
        	
        	if (component == null) {
        		project.VBComponents.Import(f.FullName);
        	} else {
        		if (component.Type == vbext_ComponentType.vbext_ct_Document) {
        			var tempComp = project.VBComponents.Import(f.FullName);
        			component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines);
        			var S = tempComp.CodeModule.Lines[1, tempComp.CodeModule.CountOfLines];
        			component.CodeModule.InsertLines(1, S);
        			project.VBComponents.Remove(tempComp);
        		}
        	}
        	f.Delete();
        }
    }
}
