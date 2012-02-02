using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

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
					bool rukGr=isNeed(wsInstr, rowIndex, 14);
					bool gidr=isNeed(wsInstr, rowIndex, 15);
					bool injRas=isNeed(wsInstr, rowIndex, 16);
					bool teh=isNeed(wsInstr, rowIndex, 17);
					bool create=nss || dem || mg || dgshu || inj || rukGr || gidr || injRas || teh;
					if (create) {
						string name=wsInstr.Cells[3][rowIndex].Value.ToString();
						Logger.log(name);
						name = String.Format("С документом \n \"{0}\"", name);
						wsList.Copy(newWS);
						Worksheet ws=newWB.Worksheets["Лист ознакомления"];
						ws.Name = String.Format("{0}-{1}", folder, instr);
						ws.Cells[1][2].Value = name;

						Dictionary<string,string> data=getData(ws);
						if (!nss) {
							removeDoljn(data, "Нач. ОС");
							removeDoljn(data, "Зам.нач.ОС");
							removeDoljn(data, "НСС");
							removeDoljn(data, "НС");							
						}
						if (!dem) {
							removeDoljn(data, "ДЭМ");
						}
						if (!dgshu) {
							removeDoljn(data, "ДГЩУ");
						}
						if (!mg) {
							removeDoljn(data, "МГ");
						}
						if (!inj) {
							removeDoljn(data, "инженер");
						}
						if (!rukGr) {
							removeDoljn(data, "Рук. гр. режимов");
						}
						if (!injRas) {
							removeDoljn(data, "инженер по расчетам");
						}
						if (!gidr) {
							removeDoljn(data, "инженер-гидролог");
						}
						if (!teh) {
							removeDoljn(data, "техник");
						}
						refreshList(ws,data);
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

		Dictionary<string, string> getData(Worksheet sheet) {
			Dictionary<string,string> data=new Dictionary<string, string>();
			for (int rowIndex=5; rowIndex <= oznakomLastString; rowIndex++) {
				object val=sheet.Cells[2][rowIndex].Value;
				string strVal=val == null ? "" : val.ToString();

				object key=sheet.Cells[3][rowIndex].Value;
				string strKey=key == null ? "" : key.ToString();

				if (!String.IsNullOrEmpty(strKey)) {
					data.Add(strKey, strVal);
				}

				sheet.Cells[2][rowIndex].Value = "";
				sheet.Cells[3][rowIndex].Value = "";
			}
			return data;
		}

		protected void removeDoljn(Dictionary<string, string> data, string doljn) {
			List<string> forRemove=new List<string>();
			foreach (KeyValuePair<string,string> de in data) {
				if (de.Value.ToLower().Trim().Equals(doljn.ToLower())) {
					forRemove.Add(de.Key);
				}

			}

			foreach (string key in forRemove) {
				data.Remove(key);
			}

		}

		protected void refreshList(Worksheet sheet, Dictionary<string,string> data) {
			int row=5;
			foreach (KeyValuePair<string,string> de in data) {
				sheet.Cells[3][row].Value = de.Key;
				sheet.Cells[2][row].Value = de.Value;
				row++;
			}

			sheet.Range[sheet.Cells[2][5], sheet.Cells[3][oznakomLastString]].Borders.LineStyle = XlLineStyle.xlContinuous;

		}

	}

	
}
