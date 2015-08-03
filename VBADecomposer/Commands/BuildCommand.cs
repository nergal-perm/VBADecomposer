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
		private string _workbookPath;
		private string _sourceFolder;
		
		public BuildCommand(string[] commandLine)
			: base(commandLine) {
		}
	
		#region Implemented abstract members

		public override bool run() {
			ImportCode();
			return true;
		}

		public override bool argsAreOk() {
			bool fileParam = false; //-file parameter is required, so we keep a flag to show if it's present
			bool folderParam = false;
			int argsCount = _commandLine.Length;

			for (int i = 1; i < argsCount; i++) {
				switch (_commandLine[i].ToUpper()) {
					case "-FILE":
						if (i + 1 < argsCount) {
							_workbookPath = _commandLine[i + 1];
							fileParam = true;
						} else {
							return false;
						}
						break;
					case "-FOLDER":
						if (i + 1 < argsCount) {
							_sourceFolder = _commandLine[i + 1];
							folderParam = true;
						} else {
							return false;
						}
						break;
					default:
						break;
				}
			}        	
			return fileParam && folderParam;
		}

		#endregion
        
		public void ImportCode() {
			Excel.Application _xlApp = new Excel.Application();
			Excel.Workbook wb = _xlApp.Workbooks.Add();
        	
			foreach (FileInfo file in new DirectoryInfo(_sourceFolder).GetFiles()) {
				if (file.Extension == ".guid") {
					// TODO: Implement references import
					//ImportReferences(file, wb);
				} else {
					ImportComponent(file, wb);
				}
			}
        	
			wb.SaveAs(_workbookPath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
			wb.Close(false);
			Marshal.ReleaseComObject(wb);
			_xlApp.Quit();        	
		}
        
		private void ImportReferences(FileInfo f, Excel.Workbook wb) {
			VBProject project = wb.VBProject;
			var sr = f.OpenText();
			while (!sr.EndOfStream) {
				string referencePath = sr.ReadLine();
				Console.WriteLine(referencePath);
				project.References.AddFromFile(sr.ReadLine());
			}
		}
		
        
		private void ImportComponent(FileInfo f, Excel.Workbook wb) {
			VBA.VBProject project = wb.VBProject;
			VBA.VBComponent component = null;
			string moduleName = f.Name.Substring(0, f.Name.LastIndexOf("."[0]));
			try {
				Console.Write("Компонент {0}. ", moduleName);
				component = project.VBComponents.Item(moduleName);
			} catch (Exception e) {
				if (f.Extension ==".wks") {
//					Console.Write("Добавляем новый лист, было {0}", wb.Worksheets.Count);
//					Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.Add();
//					Console.WriteLine(", стало {0}", wb.Worksheets.Count);
//					ws._CodeName = moduleName;
//					component = project.VBComponents.Add(vbext_ComponentType.vbext_ct_Document);
//					component.Name = moduleName;
				}
			}
        	
			if (component == null) {
				Console.WriteLine("Импорт нового модуля из файла {0}", f.FullName);
				project.VBComponents.Import(f.FullName);
			} else {
				if (component.Type == vbext_ComponentType.vbext_ct_Document) {
					Console.WriteLine("Импорт существующего модуля: {0}", component.Name);	
					var tempComp = project.VBComponents.Import(f.FullName);
					component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines);
					if (tempComp.CodeModule.CountOfLines > 0) {
						var S = tempComp.CodeModule.Lines[1, tempComp.CodeModule.CountOfLines];
						component.CodeModule.InsertLines(1, S);
					}
					project.VBComponents.Remove(tempComp);
				}
			}
			f.Delete();
		}
	}
}
