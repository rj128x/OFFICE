using System;
using Microsoft.Office.Interop.Word;

namespace Office.Shared
{
	public enum BlankOperation{replace,tip,current}
	class ProcessBlanks
	{
		protected FolderOperations folder=null;
		protected Application app=new Application();
		protected Document doc;
		public ProcessBlanks(string path, BlankOperation oper,bool Visible) {
			app.Visible = Visible;
			FolderOperations folder=new FolderOperations();
			switch (oper) {
				case BlankOperation.replace:
					folder.processFile = replaceNSGES;
					break;
				case BlankOperation.current:
					folder.processFile = createBlank;
					break;
				case BlankOperation.tip:
					folder.processFile = createTipBlank;
					break;
			}
			folder.calcFolder(path);			
		}

		protected bool replaceNSGES(string fileName) {			
			doc = app.Documents.Open(fileName, Visible: true, ReadOnly: false);

			int count=0;
			while (doc.Range().Find.Execute(FindText: "НС ГЭС", MatchCase: false, ReplaceWith: "НСС", Replace: WdReplace.wdReplaceOne)) {
				count++;
			}
			Logger.log(fileName+"-"+count.ToString());
			
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
				doc.Tables[doc.Tables.Count - 1].Rows.Last.Range.ParagraphFormat.KeepWithNext = -1;
				doc.Tables[doc.Tables.Count - 1].Rows.Last.Range.ParagraphFormat.KeepTogether = -1;
			}

			Paragraph last=doc.Paragraphs.Last;
			while (cntNotNull <= cnt) {
				if (last.Range.Text.Trim().Length > 3) {
					//Logger.log(last.Range.Text);
					cntNotNull++;
				}
				last.Format.KeepWithNext = -1;
				last.Format.KeepTogether = -1;
				last = last.Previous();
			}
			
		}

		protected bool createTipBlank(string fileName) {
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

				
				first.Rows[1].Cells[1].Range.Text="Типовой бланк переключений\n №" + getNumber(fileName);
				first.Rows[1].Cells[2].Range.Text = "Утверждаю\nГлавный инженер филиала\nОАО \"РусГидро\" - \"Воткинская ГЭС\"\n__________________А.П.Деев\n\"____\"_____________2012г.";
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
				last.Rows[2].Cells[3].Range.Text = "Цирлин С.Л.";

				last.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				last.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				last.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				last.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				last.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

				//replaceSpaces(first);
				//replaceSpaces(last);

				replaceSpaces(doc);
				addFooter(fileName, doc,true);
				podpis(doc,7);
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
				
				Table tab =first.Rows[1].Cells[1].Range.Tables.Add(first.Rows[1].Cells[1].Range, 1, 3);
				tab.Rows.Add();
				tab.Rows.Add();
							
				

				tab.Rows[1].Cells[1].Merge(tab.Rows[1].Cells[2]);
				tab.Rows[1].Cells[1].Merge(tab.Rows[1].Cells[2]);
				tab.Rows[1].Cells[1].Range.Text = "Бланк переключений";
								
				
				tab.Rows[3].Cells[1].Merge(tab.Rows[3].Cells[2]);
				tab.Rows[3].Cells[1].Merge(tab.Rows[3].Cells[2]);
				tab.Rows[3].Cells[1].Range.Text="Начало______час______мин\nКонец______час______мин";

				tab.Rows[2].Cells[3].Range.Paragraphs.Add();
				tab.Rows[2].Cells[3].Range.Paragraphs.First.Range.Select();
				app.Selection.Range.Fields.Add(app.Selection.Range, Text: "Date ");
				tab.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				tab.Rows[2].Cells[3].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

				tab.Rows[2].Cells[1].Range.Text = "№     ";
				tab.Rows[2].Cells[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
				tab.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;


				tab.Rows[2].Cells[2].Range.Text = "/" + getNumber(fileName);
				tab.Rows[2].Cells[2].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
				tab.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

				tab.Rows[2].Cells[1].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
				tab.Rows[2].Cells[2].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
				tab.Rows[2].Cells[3].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;

				tab.Rows[2].Cells[1].PreferredWidth = 20;
				tab.Rows[2].Cells[2].PreferredWidth = 50;
				tab.Rows[2].Cells[3].PreferredWidth = 30;

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
				addFooter(fileName, doc,false);
				podpis(doc,10);
				
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


		protected void replaceSpaces(Document doc) {
			int count=0;
			doc.Range().Find.Execute(FindText: "^w", MatchCase: false, ReplaceWith: " ", Replace: WdReplace.wdReplaceAll) ;
			doc.Range().Find.Execute(FindText: "^w^p", MatchCase: false, ReplaceWith: "^p", Replace: WdReplace.wdReplaceAll);
			for (int i=0;i<20;i++){
				doc.Range().Find.Execute(FindText: "^p^p", MatchCase: false, ReplaceWith: "^p", Replace: WdReplace.wdReplaceAll);
			}
		}

		protected string getNumber(string fileName) {
			string number="";

			try {
				string[] fns=fileName.Split("\\".ToCharArray());
				string fn=fns[fns.Length - 1];
				char[] separ= { ' ', '-', };
				fns = fn.Split(separ);
				fn = fns[0];
				number = fns[0];
			} catch {
			}
			return number;
		}



		protected void addFooter(string fileName,Document doc, bool addTip) {
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
