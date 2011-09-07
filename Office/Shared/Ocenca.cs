using System;
using Microsoft.Office.Interop.Excel;

namespace Office.Shared
{
	class Ocenca
	{
		protected string fileName;
		protected int oznakomLastString=49;
		protected Application app=new Application();
		protected bool visible;
		public Ocenca(string fileName, bool visible) {
			this.fileName = fileName;
			this.visible = visible;
		}


		public void processFile() {
			app.Visible = visible;
			Workbook wb=app.Workbooks.Open(fileName, ReadOnly: true);
			Workbook newWB=app.Workbooks.Add();
			Worksheet wsPeople=wb.Worksheets["Список"];
			Worksheet wsBlank;
			Worksheet newWS=newWB.Worksheets.Add();


			for (int rowIndex=2; rowIndex <= 41; rowIndex++) {
				string name=wsPeople.Cells[1][rowIndex].Value.ToString();
				string vvod = DateTime.Parse(wsPeople.Cells[7][rowIndex].Value.ToString()).ToShortDateString();
				string perv = DateTime.Parse(wsPeople.Cells[8][rowIndex].Value.ToString()).ToShortDateString();
				string staz = wsPeople.Cells[9][rowIndex].Value.ToString().Length>1?DateTime.Parse(wsPeople.Cells[9][rowIndex].Value.ToString()).ToShortDateString():"";
				string dubl = wsPeople.Cells[10][rowIndex].Value.ToString().Length>1?DateTime.Parse(wsPeople.Cells[10][rowIndex].Value.ToString()).ToShortDateString():"";
				string kind = wsPeople.Cells[11][rowIndex].Value.ToString();
				string vuz = wsPeople.Cells[12][rowIndex].Value.ToString();
				string dipl = wsPeople.Cells[13][rowIndex].Value.ToString();
				string udost = wsPeople.Cells[14][rowIndex].Value.ToString();
				string udostDat = udost.Length>1?DateTime.Parse(wsPeople.Cells[15][rowIndex].Value.ToString()).ToShortDateString():"";
				string udost1 = wsPeople.Cells[16][rowIndex].Value.ToString();
				string udostDat1 = udost1.Length > 1 ? DateTime.Parse(wsPeople.Cells[17][rowIndex].Value.ToString()).ToShortDateString() : "";
				string prov = wsPeople.Cells[18][rowIndex].Value.ToString().Length>1?DateTime.Parse(wsPeople.Cells[18][rowIndex].Value.ToString()).ToShortDateString():"";
				string povt = wsPeople.Cells[19][rowIndex].Value.ToString();

				string[] names = name.Split(' ');
				string shortName= names[0] + ' ' + names[1].Substring(0, 1) + ' ' + names[2].Substring(0, 1);

				string blank = wsPeople.Cells[2][rowIndex].Value;
				Logger.log(name);

				wsBlank = wb.Worksheets["бланк_" + blank];

				wsBlank.Copy(newWS);
				Worksheet ws = newWB.Worksheets["бланк_" + blank];
				int len = shortName.Length < 30 ? shortName.Length : 30;
				ws.Name = String.Format("{0}", shortName.Substring(0, len));
				ws.Cells[3][3].Value = name;
				ws.Cells[1][8].Value = "Оперативная служба";
				ws.Cells[4][19].Value = vvod;
				ws.Cells[4][20].Value = perv;
				ws.Cells[4][23].Value = prov;
				if ((blank == "ДЭМ") || (blank == "МГ") || (blank == "НС") || (blank == "НСС") || (blank == "ДГЩУ")) {
					ws.Cells[4][24].Value = staz;
					ws.Cells[4][25].Value = dubl;
				}
				if (dipl.Length < 2) {
					ws.Cells[1][30].Value += String.Format("\n     {0}", kind);
				} else {
					ws.Cells[1][30].Value += String.Format("\n     {0}: {1} Диплом № {2}", kind, vuz, dipl);
				}
				if (udost.Length > 1) {
					ws.Cells[1][38].Value = String.Format("{0}\n     Удостоверение № {1} от {2}", "Удостоверение о проверке знаний", udost, udostDat);
				} else {
					ws.Cells[1][38].Value = String.Format("{0}\n", "Удостоверение о проверке знаний");
				}
				
				if (udost1.Length > 1) {
					ws.Cells[1][38].Value += String.Format("\n{0}\n     Удостоверение № {1} от {2}", "Удостоверение на право обслуживания объектов ГосГорТехНадзора", udost1, udostDat1);
                }
                else if (((blank == "ДЭМ") || (blank == "МГ") || (blank == "НС") || (blank == "НСС") || (blank == "ЗНОС")))
                {
                    ws.Cells[1][38].Value += String.Format("\n{0}", "Удостоверение на право обслуживания объектов ГосГорТехНадзора");
				}

				if (povt.Length > 1) {
					string[] povts=povt.Split(' ');
					int index=0;
					foreach (string p in povts) {
						ws.Cells[4 + index][27].Value = p;
						index++;
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
			for (int rowIndex=5; rowIndex <= 46; rowIndex++) {
				object val=sheet.Cells[2][rowIndex].Value;
				string strVal=val == null ? "" : val.ToString();
				if (strVal.ToLower().Trim().Equals(doljn.ToLower())) {
					sheet.Cells[2][rowIndex].Value = "";
					sheet.Cells[3][rowIndex].Value = "";
				}
			}
		}
	}


}
