using System;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace Office.Shared
{
	class ProcessPDF
	{
		protected FolderOperations folder=null;
		protected Application app=new Application();

		protected bool onlyFirstPage;
		protected string pathPDF;
		protected string path;

		protected Document doc;
		public ProcessPDF(string path, string pathPDF, bool onlyFirstPage, bool Visible) {
			this.onlyFirstPage = onlyFirstPage;
			this.path = path;
			this.pathPDF = pathPDF;
			app.Visible = Visible;
			FolderOperations folder=new FolderOperations();
			folder.processFile = Operation;

			folder.calcFolder(path);			
		}

		protected bool Operation(string fileName) {
			FileInfo fileInfo=new FileInfo(fileName);
			Logger.log(fileName);
			try{
					string newFileName=fileName.Replace(path, pathPDF).Replace(".docx", ".pdf").Replace(".doc", ".pdf");
					string dir=fileInfo.Directory.FullName.Replace(path, pathPDF);
					Directory.CreateDirectory(dir);
					if (File.Exists(newFileName)) {
						File.Delete(newFileName);
					}
					createPDF(fileName, true, newFileName);
					Logger.log(newFileName);
					return true;				
				
			}catch (Exception e){
				Logger.log("ERROR "+e.ToString());
			}
			return true;

		}

		
		protected void saveAsPDF(Document wordDocument, string fileTo){
			string paramExportFilePath = fileTo;
			WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
			bool paramOpenAfterExport = false;
			WdExportOptimizeFor paramExportOptimizeFor =
				 WdExportOptimizeFor.wdExportOptimizeForPrint;
			WdExportRange paramExportRange = onlyFirstPage? WdExportRange.wdExportFromTo:WdExportRange.wdExportAllDocument;
			int paramStartPage = 0;
			int paramEndPage = 0;
			if (onlyFirstPage) {
				paramStartPage = 1;
				paramEndPage = 1;
			} 
			WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
			bool paramIncludeDocProps = true;
			bool paramKeepIRM = true;
			WdExportCreateBookmarks paramCreateBookmarks = 
				 WdExportCreateBookmarks.wdExportCreateWordBookmarks;
			bool paramDocStructureTags = true;
			bool paramBitmapMissingFonts = true;
			bool paramUseISO19005_1 = false;
			object paramMissing = Type.Missing;

			if (wordDocument != null)
				wordDocument.ExportAsFixedFormat(paramExportFilePath,
            paramExportFormat, paramOpenAfterExport, 
            paramExportOptimizeFor, paramExportRange, paramStartPage,
            paramEndPage, paramExportItem, paramIncludeDocProps, 
            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags, 
            paramBitmapMissingFonts, paramUseISO19005_1,
            ref paramMissing);
		}

		protected bool createPDF(string fileName, bool pdf, string pdfName) {
			doc = null;
			try {
				doc = app.Documents.Open(fileName, Visible: true, ReadOnly: false);
				doc.Select();

				if (pdf) {
					saveAsPDF(doc, pdfName);
				}
				Logger.log(fileName);
				(doc as _Document).Close(SaveChanges: false);
				return true;
			} catch (Exception e) {
				try { (doc as _Document).Close(SaveChanges: false); } catch { }
				Logger.log("ERROR: " + fileName);
				Logger.log("--" + e.Message);
				return false;
			}
		}
	}
}
