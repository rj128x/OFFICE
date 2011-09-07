using System;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Office.Shared
{
	class InstrFolders
	{
		protected string fileName;
		protected string outPath;


		protected Application app=new Application();
		protected bool visible;
		public InstrFolders(string fileName, string outPath, bool visible) {
			this.fileName = fileName;
			this.outPath = outPath;
			this.visible = visible;
		}


		public void processFile() {
			app.Visible = visible;
			Workbook wb=app.Workbooks.Open(fileName, ReadOnly: true);
			Workbook newWB=app.Workbooks.Add();
			Worksheet wsInstr=wb.Worksheets["Список инструкций"];
			DirectoryInfo outDir=new DirectoryInfo(outPath);


			int prevFolder=-1;
			DirectoryInfo folderDir=null;
			for (int rowIndex=1; rowIndex <= 250; rowIndex++) {
				int folder;
				int instr;
				try {
					object folderVal=wsInstr.Cells[1][rowIndex].Value;
					object instrVal=wsInstr.Cells[2][rowIndex].Value;
					Int32.TryParse(folderVal.ToString(), out folder);
					Int32.TryParse(instrVal.ToString(), out instr);
				} catch {
					folder = 0;
					instr = 0;
				}
				if ((folder > 0) && (instr > 0)) {
					if (folder != prevFolder) {
						folderDir = outDir.CreateSubdirectory(folder.ToString());
					}
					string name=wsInstr.Cells[3][rowIndex].Value.ToString();
					string fname=String.Format("[{1}] {0}", name, instr);
					fname=fname.Replace('\"',' ');
					fname = fname.Replace('\\', '-');
					fname = fname.Replace('/', '-');
					fname = fname.Replace('.', ' ');
					fname = fname.Replace(';', ' ');
					fname = fname.Replace(':', ' ');
					fname = fname.Replace('\n', ' ');
					int len=folderDir.FullName.Length;
					int maxLen=240 - len;
					int newLen=fname.Length > maxLen ? maxLen : fname.Length;
					fname = fname.Substring(0, newLen);
					folderDir.CreateSubdirectory(fname);
				}
				System.Windows.Forms.Application.DoEvents();
			}
			app.Visible = true;
		}


	}

	

	
}
