using System;
using Microsoft.Office.Interop.Excel;

namespace Office.Shared
{
	class ListsZnak
	{
		protected string fileName;
		protected int oznakomLastString=49;
		protected Application app=new Application();
		protected bool visible;
		public ListsZnak(string fileName, bool visible) {
			this.fileName = fileName;
			this.visible = visible;
		}


		public void processFile() {
			app.Visible = visible;
			Workbook wb=app.Workbooks.Open(fileName, ReadOnly: true);
			Workbook newWB=app.Workbooks.Add();
			Worksheet wsInstr=wb.Worksheets["Список инструкций"];
			Worksheet wsList=wb.Worksheets["Лист ознакомления"];
			Worksheet newWS=newWB.Worksheets.Add();


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
					bool nss=isNeed(wsInstr, rowIndex, 9);
					bool dem=isNeed(wsInstr, rowIndex, 10);
					bool mg=isNeed(wsInstr, rowIndex, 11);
					bool dgshu=isNeed(wsInstr, rowIndex, 12);
					bool inj=isNeed(wsInstr, rowIndex, 13);
					bool other=isNeed(wsInstr, rowIndex, 14);
					bool create=nss || dem || mg || dgshu || inj || other;
					if (create) {
						string name=wsInstr.Cells[3][rowIndex].Value.ToString();
						Logger.log(name);
						name = String.Format("С документом \n \"{0}\"", name);
						wsList.Copy(newWS);
						Worksheet ws=newWB.Worksheets["Лист ознакомления"];
						ws.Name = String.Format("{0}-{1}", folder, instr);
						ws.Cells[1][2].Value = name;
						if (!nss) {
							removeDoljn(ws, "Нач. ОС");
							removeDoljn(ws, "Зам.нач.ОС");
							removeDoljn(ws, "НСС");
							removeDoljn(ws, "НС");
						}
						if (!dem) {
							removeDoljn(ws, "ДЭМ");
						}
						if (!dgshu) {
							removeDoljn(ws, "ДГЩУ");
						}
						if (!mg) {
							removeDoljn(ws, "МГ");
						}
						if (!inj) {
							removeDoljn(ws, "Инженер");
						}
						if (!other) {
							removeDoljn(ws, "техник");
							removeDoljn(ws, "гидролог");
							removeDoljn(ws, "специалист");
						}
						refreshList(ws);
					}
				}
				System.Windows.Forms.Application.DoEvents();
			}
			newWS.Delete();
			app.Visible = true;
		}
		protected bool isNeed(Worksheet sheet, int row, int col) {
			object val=sheet.Cells[col][row].Value;
			return val == null ? false : val.ToString().ToUpper().Equals("V");
		}

		protected void removeDoljn(Worksheet sheet, string doljn) {
			for (int rowIndex=5; rowIndex <= 49; rowIndex++) {
				object val=sheet.Cells[2][rowIndex].Value;
				string strVal=val == null ? "" : val.ToString();
				if (strVal.ToLower().Trim().Equals(doljn.ToLower())) {
					sheet.Cells[2][rowIndex].Value = "";
					sheet.Cells[3][rowIndex].Value = "";
				}
			}

		}

		protected void refreshList(Worksheet sheet) {
			for (int rowIndex=5; rowIndex <= oznakomLastString; rowIndex++) {
				object val=sheet.Cells[2][rowIndex].Value;
				string strVal=val == null ? "" : val.ToString();
				if (strVal.Length == 0) {
					for (int ri=rowIndex + 1; ri <= oznakomLastString; ri++) {
						object nextVal=sheet.Cells[2][ri].Value;
						string strNextVal=nextVal == null ? "" : nextVal.ToString();
						if (strNextVal.Length > 0) {
							sheet.Range[sheet.Cells[2][ri], sheet.Cells[3][oznakomLastString]].Cut(sheet.Cells[2][rowIndex]);
							rowIndex = ri;
							break;
						}
					}

				}
			}
			sheet.Range[sheet.Cells[2][5], sheet.Cells[3][oznakomLastString]].Borders.LineStyle = XlLineStyle.xlContinuous;

		}

	}

	
}
