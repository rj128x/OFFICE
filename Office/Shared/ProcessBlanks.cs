using System;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;

namespace Office.Shared
{
	public enum BlankOperation { replace, tip, current }
	class ProcessBlanks
	{
		protected FolderOperations folder=null;
		protected Application app=new Application();

		protected bool createPDF;
		protected bool createTip;
		protected bool createCurrent;
		protected bool onlyFirstPage;
		protected string pathTip;
		protected string pathPDF;
		protected string pathCurrent;
		protected string path;

		protected Document doc;
		public ProcessBlanks(string path, string pathPDF, string pathTip, string pathCurrent, bool createPDF, bool onlyFirstPage, bool createTip, bool createCurrent, bool Visible) {
			this.createCurrent = createCurrent;
			this.createPDF = createPDF;
			this.createTip = createTip;
			this.onlyFirstPage = onlyFirstPage;
			this.path = path;
			this.pathCurrent = pathCurrent;
			this.pathPDF = pathPDF;
			this.pathTip = pathTip;
			app.Visible = Visible;
			FolderOperations folder=new FolderOperations();
			folder.processFile = Operation;

			folder.calcFolder(path);
		}

		protected bool Operation(string fileName) {
			FileInfo fileInfo=new FileInfo(fileName);
			Logger.log(fileName);
			try {
				if (createTip) {
					string newFileName=fileName.Replace(path, pathTip);
					string dir=fileInfo.Directory.FullName.Replace(path, pathTip);
					Directory.CreateDirectory(dir);
					string newFileNamePDF=fileName.Replace(path, pathPDF).Replace(".docx", ".pdf").Replace(".doc", ".pdf");
					dir = fileInfo.Directory.FullName.Replace(path, pathPDF);
					Directory.CreateDirectory(dir);
					if (File.Exists(newFileName)) {
						File.Delete(newFileName);
					}
					if (File.Exists(newFileNamePDF)) {
						File.Delete(newFileNamePDF);
					}
					Logger.log(newFileName);

					File.Copy(fileName, newFileName);
					createTipBlank(newFileName, createPDF, newFileNamePDF);
					return true;
				}
				if (createCurrent) {
					string newFileName=fileName.Replace(path, pathCurrent);
					string dir=fileInfo.Directory.FullName.Replace(path, pathTip);
					Directory.CreateDirectory(dir);
					if (File.Exists(newFileName)) {
						File.Delete(newFileName);
					}
					Logger.log(newFileName);
					File.Copy(fileName, newFileName);
					createBlank(newFileName);
					return true;
				}
			} catch (Exception e) {
				Logger.log("ERROR " + e.ToString());
			}
			return true;

		}

		protected bool replaceNSGES(string fileName) {
			doc = app.Documents.Open(fileName, Visible: true, ReadOnly: false);

			int count=0;
			while (doc.Range().Find.Execute(FindText: "НС ГЭС", MatchCase: false, ReplaceWith: "НСС", Replace: WdReplace.wdReplaceOne)) {
				count++;
			}
			Logger.log(fileName + "-" + count.ToString());

			(doc as _Document).Close(SaveChanges: true);

			return true;
		}

		protected void podpis(Document doc, int cnt) {
			int count=doc.Paragraphs.Count;
			int cntNotNull=0;

			//doc.Select();
			//doc.SelectAllEditableRanges();
			doc.Paragraphs.Format.KeepTogether = 0;
			doc.Paragraphs.Format.KeepWithNext = 0;

			if (doc.Tables.Count > 2) {
				try {
					doc.Tables[doc.Tables.Count - 1].Rows.Last.Range.ParagraphFormat.KeepWithNext = -1;
					doc.Tables[doc.Tables.Count - 1].Rows.Last.Range.ParagraphFormat.KeepTogether = -1;
				}catch{}
			}

			Paragraph last=doc.Paragraphs.Last;
			while (cntNotNull <= cnt) {
				if (last.Range.Text.Trim().Length > 3) {
					//Logger.log(last.Range.Text);
					cntNotNull++;
				}
				try {
					last.Format.KeepWithNext = -1;
					last.Format.KeepTogether = -1;
				} catch { }
				last = last.Previous();
			}

		}

		protected void saveAsPDF(Document wordDocument, string fileTo) {
			string paramExportFilePath = fileTo;
			WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
			bool paramOpenAfterExport = false;
			WdExportOptimizeFor paramExportOptimizeFor =	 WdExportOptimizeFor.wdExportOptimizeForPrint;
			WdExportRange paramExportRange = onlyFirstPage ? WdExportRange.wdExportFromTo : WdExportRange.wdExportAllDocument;
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
				 WdExportCreateBookmarks.wdExportCreateNoBookmarks;
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

		protected bool createTipBlank(string fileName, bool pdf, string pdfName) {
			doc = null;
			try {
				doc = app.Documents.Open(fileName, Visible: true, ReadOnly: false);
				doc.Select();

				doc.PageSetup.LeftMargin = 60;
				doc.PageSetup.TopMargin = 30;
				doc.PageSetup.BottomMargin = 30;
				doc.PageSetup.RightMargin = 30;
				doc.PageSetup.FooterDistance = 30;
				doc.PageSetup.HeaderDistance = 0;

				int x=doc.Range().Tables.Count;

				Table first=doc.Range().Tables[1];
				resetTable(first);
				first.Range.Font.Size = 13;
				first.Columns.Add();
				first.AutoFormat(ApplyBorders: false, ApplyShading: false, ApplyFont: false, ApplyColor: false, ApplyHeadingRows: false,
				ApplyLastRow: false, ApplyFirstColumn: false, AutoFit: true, ApplyLastColumn: false);

				first.Rows[1].Cells[1].Range.Text = "Типовой бланк переключений \n№" + getNumber(fileName);
				first.Rows[1].Cells[2].Range.Text = "Утверждаю\nГлавный инженер филиала\nОАО \"РусГидро\" - \"Воткинская ГЭС\"\n__________________Э.М. Скрипка\n\"____\"_____________2012г.";
				first.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);


				Table last=doc.Range().Tables[doc.Range().Tables.Count];
				resetTable(last);
				last.Range.Font.Size = 12;
				last.TopPadding = 20;
				last.Columns.Add();
				last.Columns.Add();
				last.Rows.Add();
				last.AutoFormat(ApplyBorders: false, ApplyShading: false, ApplyFont: false, ApplyColor: false, ApplyHeadingRows: false,
				ApplyLastRow: false, ApplyFirstColumn: false, AutoFit: true, ApplyLastColumn: false);

				last.Rows[1].Cells[1].Range.Text = "Начальник СТСУ";
				last.Rows[1].Cells[2].Range.Text = "_____________________";
				last.Rows[1].Cells[3].Range.Text = "Кочеев Н.Н.";
				last.Rows[2].Cells[1].Range.Text = "Начальник ОС";
				last.Rows[2].Cells[2].Range.Text = "_____________________";
				last.Rows[2].Cells[3].Range.Text = "Иванов А.В.";

				last.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				last.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				last.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				last.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				last.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

				//replaceSpaces(first);
				//replaceSpaces(last);

				replaceSpaces(doc);
				makeSchema(doc);
				addFooter(fileName, doc, true);
				podpis(doc, 7);

				if (pdf) {
					saveAsPDF(doc, pdfName);
				}
				Logger.log(fileName);
				(doc as _Document).Close(SaveChanges: true);
				return true;
			} catch (Exception e) {
				try { (doc as _Document).Close(SaveChanges: false); } catch { }
				Logger.log("ERROR: " + fileName);
				Logger.log("--" + e.Message);
				return false;
			}
		}


		protected bool createBlank(string fileName) {
			try {
				doc = app.Documents.Open(fileName, Visible: true, ReadOnly: false);
				doc.Select();

				doc.PageSetup.LeftMargin = 60;
				doc.PageSetup.TopMargin = 30;
				doc.PageSetup.BottomMargin = 30;
				doc.PageSetup.RightMargin = 30;
				doc.PageSetup.FooterDistance = 30;
				doc.PageSetup.HeaderDistance = 0;

				int x=doc.Range().Tables.Count;

				Table first=doc.Range().Tables[1];
				resetTable(first);
				first.Range.Font.Size = 13;
				first.Columns.Add();
				first.AutoFormat(ApplyBorders: false, ApplyShading: false, ApplyFont: false, ApplyColor: false, ApplyHeadingRows: false,
				ApplyLastRow: false, ApplyFirstColumn: false, AutoFit: false, ApplyLastColumn: false);
				first.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

				Table tab =first.Rows[1].Cells[1].Range.Tables.Add(first.Rows[1].Cells[1].Range, 1, 4);
				tab.Rows.Add();
				tab.Rows.Add();
				tab.Rows.Add();


				tab.Rows[1].Cells[1].Merge(tab.Rows[1].Cells[2]);
				tab.Rows[1].Cells[1].Merge(tab.Rows[1].Cells[2]);
				tab.Rows[1].Cells[1].Merge(tab.Rows[1].Cells[2]);
				tab.Rows[1].Cells[1].Range.Text = "Бланк переключений";



				tab.Rows[2].Cells[4].Range.Paragraphs.Add();
				tab.Rows[2].Cells[4].Range.Paragraphs.First.Range.Select();
				app.Selection.Range.Fields.Add(app.Selection.Range, Text: "Date ");
				tab.Rows[2].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

				tab.Rows[2].Cells[1].Range.Text = "№";
				tab.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

				tab.Rows[2].Cells[2].Range.Text = "______";
				tab.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				tab.Rows[2].Cells[2].Range.Font.Color = WdColor.wdColorWhite;


				tab.Rows[2].Cells[3].Range.Text = "/" + getNumber(fileName);
				tab.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

				tab.Rows[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

				tab.Rows[2].Cells[1].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto;
				tab.Rows[2].Cells[2].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto;
				tab.Rows[2].Cells[3].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto;
				tab.Rows[2].Cells[4].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto;


				tab.Rows[3].Cells[1].Merge(tab.Rows[3].Cells[2]);
				tab.Rows[3].Cells[1].Merge(tab.Rows[3].Cells[2]);
				tab.Rows[3].Cells[1].Merge(tab.Rows[3].Cells[2]);
				tab.Rows[3].Cells[1].Range.Text = "Начало ____час____мин";
				tab.Rows[3].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

				tab.Rows[4].Cells[1].Merge(tab.Rows[4].Cells[2]);
				tab.Rows[4].Cells[1].Merge(tab.Rows[4].Cells[2]);
				tab.Rows[4].Cells[1].Merge(tab.Rows[4].Cells[2]);
				tab.Rows[4].Cells[1].Range.Text = "Конец ____час____мин";
				tab.Rows[4].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


				/*tab.Rows[2].Cells[1].PreferredWidth = 20;
				tab.Rows[2].Cells[2].PreferredWidth = 50;
				tab.Rows[2].Cells[3].PreferredWidth = 30;*/

				first.Rows[1].Cells[2].Range.InsertAfter("Филиал ОАО \"РусГидро\"\n - \"Воткинская ГЭС\"");

				first.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

				Table last=doc.Range().Tables[doc.Range().Tables.Count];
				resetTable(last);
				last.Range.Font.Size = 12;
				last.TopPadding = 10;
				last.Columns.Add();
				last.Rows.Add();
				last.Rows.Add();
				last.Rows.Add();

				last.AutoFormat(ApplyBorders: false, ApplyShading: false, ApplyFont: false, ApplyColor: false, ApplyHeadingRows: false,
				ApplyLastRow: false, ApplyFirstColumn: false, AutoFit: false, ApplyLastColumn: false);


				last.Rows[1].Cells[1].Range.Text = "Типовой бланк переключений проверен, соответствует схемам, переключения в указанной в нем последовательности могут быть выполнены";
				last.Rows[2].Cells[1].Range.Text = "Переключения разрешаю (НСС):";
				last.Rows[2].Cells[2].Range.Text = "_________/_____________________/";
				last.Rows[3].Cells[1].Range.Text = "Лицо, производящее переключение (ДЭМ/МГ):";
				last.Rows[3].Cells[2].Range.Text = "_________/_____________________/";
				last.Rows[4].Cells[1].Range.Text = "Лицо контролируюшее (НС):";
				last.Rows[4].Cells[2].Range.Text = "_________/_____________________/";


				last.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				last.Rows[3].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				last.Rows[4].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

				last.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				last.Rows[3].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				last.Rows[4].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

				last.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
				last.AutoFormat(ApplyBorders: false, ApplyShading: false, ApplyFont: false, ApplyColor: false, ApplyHeadingRows: false,
				ApplyLastRow: false, ApplyFirstColumn: false, AutoFit: true, ApplyLastColumn: false);
				last.Rows.First.Cells.Merge();
				last.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);


				//replaceSpaces(first);
				//replaceSpaces(last);


				replaceSpaces(doc);
				addFooter(fileName, doc, false);
				podpis(doc, 10);

				Logger.log(fileName);
				(doc as _Document).Close(SaveChanges: true);
				return true;
			} catch (Exception e) {
				try { (doc as _Document).Close(SaveChanges: false); } catch { }
				Logger.log("ERROR: " + fileName);
				Logger.log("--" + e.Message);
				return false;
			}
		}


		protected void resetTable(Table table) {
			while (table.Rows.Count != 1)
				table.Rows.First.Delete();
			while (table.Columns.Count != 1)
				table.Columns.First.Delete();
			table.Rows[1].Cells[1].Range.Text = "";
			table.Range.ParagraphFormat.Reset();
			table.Range.Font.Reset();
			table.Range.Font.Name = "Times New Roman";
			table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone;
			table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;

			table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
		}

		protected void replaceSpaces(Table table) {
			table.Select();
			app.Selection.Find.Execute(FindText: " ", MatchCase: false, ReplaceWith: "^s", Replace: WdReplace.wdReplaceAll);

		}

		protected void replaceVVL(string line, string u){
			doc.Range().Find.Execute(FindText: " В "+line, MatchCase: false, ReplaceWith: " В ВЛ "+u+" "+line+" ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: " В "+u+" " + line, MatchCase: false, ReplaceWith: " В ВЛ " + u + " " + line + " ", Replace: WdReplace.wdReplaceAll);
		}



		protected void replaceSpaces(Document doc) {
			int count=0;
			doc.Range().Find.Execute(FindText: "^w", MatchCase: false, ReplaceWith: " ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "^w.", MatchCase: false, ReplaceWith: ".", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "^w^p", MatchCase: false, ReplaceWith: "^p", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "^p^w", MatchCase: false, ReplaceWith: "^p", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "–", MatchCase: false, ReplaceWith: "-", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: " -", MatchCase: false, ReplaceWith: "-", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "- ", MatchCase: false, ReplaceWith: "-", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-", MatchCase: false, ReplaceWith: "–", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "« ", MatchCase: false, ReplaceWith: "«", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: " »", MatchCase: false, ReplaceWith: "»", Replace: WdReplace.wdReplaceAll);
			
			/*doc.Range().Find.Execute(FindText: "ГЭС- ", MatchCase: false, ReplaceWith: "ГЭС -~", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-~", MatchCase: false, ReplaceWith: "- ", Replace: WdReplace.wdReplaceAll);*/

			doc.Range().Find.Execute(FindText: "-110 кВ", MatchCase: false, ReplaceWith: " 110", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-220 кВ", MatchCase: false, ReplaceWith: " 220", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-500 кВ", MatchCase: false, ReplaceWith: " 500", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-110", MatchCase: false, ReplaceWith: " 110", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-220", MatchCase: false, ReplaceWith: " 220", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "-500", MatchCase: false, ReplaceWith: " 500", Replace: WdReplace.wdReplaceAll);
			
			doc.Range().Find.Execute(FindText: "Иж-", MatchCase: false, ReplaceWith: "Ижевск ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "Иж ", MatchCase: false, ReplaceWith: "Ижевск ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "Ижевск-", MatchCase: false, ReplaceWith: "Ижевск ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "Водозабор-", MatchCase: false, ReplaceWith: "Водозабор ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "Каучук-", MatchCase: false, ReplaceWith: "Каучук ", Replace: WdReplace.wdReplaceAll);
			doc.Range().Find.Execute(FindText: "КШТ-", MatchCase: false, ReplaceWith: "КШТ ", Replace: WdReplace.wdReplaceAll);


			/*replaceVVL("КШТ 1", "110");
			replaceVVL("КШТ 2", "110");
			replaceVVL("Светлая", "110");
			replaceVVL("Ивановка", "110");
			replaceVVL("Каучук", "110");
			//replaceVVL("Светлая", "110");
			replaceVVL("ЧаТЭЦ", "110");
			replaceVVL("Березовка", "110");
			replaceVVL("Дубовая", "110");
			replaceVVL("Водозабор 1", "110");
			replaceVVL("Водозабор 2", "110");

			replaceVVL("Карманово", "500");
			replaceVVL("Емелино", "500");
			replaceVVL("Вятка", "500");

			replaceVVL("Ижевск 1", "220");
			replaceVVL("Ижевск 2", "220");
			replaceVVL("Каучук 1", "220");
			replaceVVL("Каучук 2", "220");
			//replaceVVL("Светлая", "220");*/
			for (int i=0; i < 20; i++) {
				doc.Range().Find.Execute(FindText: "^p^p", MatchCase: false, ReplaceWith: "^p", Replace: WdReplace.wdReplaceAll);
			}
		}

		protected string getObjectName(string findStr, string text) {
			int index=text.IndexOf(findStr);
			if (index > 10)
				return "";
			text = text.Remove(0,  index + findStr.Length);
			
			if (text.IndexOf("автомат") >= 0)
				return "";
			if (text.IndexOf("рубильник") >= 0)
				return "";
			text = text + ".";
			text = text.Replace(" в яч", ".");
			int i=text.IndexOf(".");
			if (i >= 0) {
				text = text.Substring(0, i);
			}
			return text;
		}

		protected String findTurn(Document doc) {
			string[] onStrs={"Отключить ", "Проверить включенное положение "};
			string[] offStrs= { "Включить ", "Проверить отключенное положение ", "Проверить отсутствие напряжения на " };
			List<String> resOn=new List<string>();
			List<String> resOff=new List<string>();			
			foreach (Paragraph p in doc.Range().Paragraphs) {
				string text=p.Range.Text;
				foreach (string onStr in onStrs) {					
					if (text.IndexOf(onStr) >= 0) {						
						string obj=getObjectName(onStr, text);
						if (obj.Length > 0) {
							if (!resOff.Contains(obj) && !resOn.Contains(obj))
								resOn.Add(obj);
						}
					}
				}

				foreach (string offStr in offStrs) {
					if (text.IndexOf(offStr) >= 0) {
						string obj=getObjectName(offStr, text);
						if (obj.Length > 0) {
							if (!resOn.Contains(obj) && !resOff.Contains(obj))
								resOff.Add(obj);
						}
					}
				}
			}
			string res="";
			if (resOn.Count > 0 && resOff.Count > 0) {
				res= "Включены: " + String.Join("; ", resOn.ToArray()) + "\nОтключены: " + String.Join("; ", resOff.ToArray()) + "\n";
			}
			return res;
		}

		protected void makeSchema(Document doc) {
			string schema=findTurn(doc);

			if (schema.Length > 0) {
				foreach (Paragraph p in doc.Range().Paragraphs) {
					string text=p.Range.Text;
					if (text.IndexOf("Исходная схема станции") == 0) {
						Paragraph newP= doc.Paragraphs.Add(p.Next().Range);
						newP.Range.Font.Bold = 0;
						newP.Range.Font.Underline = 0;
						newP.Range.Font.Italic = 1;
						newP.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
						newP.Range.Font.Size = 12;
						newP.Range.Text = schema;
						break;
					}
				}
			}
		}

		protected string getNumber(string fileName) {
			string number="";

			try {
				string[] fns=fileName.Split("\\".ToCharArray());
				string fn=fns[fns.Length - 1];
				char[] separ= { ' ', ' ', };
				fns = fn.Split(separ);
				fn = fns[0];
				number = fns[0];
			} catch {
			}
			return number;
		}



		protected void addFooter(string fileName, Document doc, bool addTip) {
			string number=getNumber(fileName);


			Range range1=doc.Sections.First.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
			range1.Text = "";
			range1.Select();
			range1.Delete();

			Range range=doc.Sections.First.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
			range.Text = "";
			range.Font.Size = 12;
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
			range.Select();

			range.Fields.Add(app.Selection.Range, Type: WdFieldType.wdFieldPage);

			if (number.Length > 0) {
				range.InsertBefore(number.ToString() + "     -");
				range.InsertAfter("-");
			}
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
		}
	}
}

