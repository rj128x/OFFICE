using System;
using Microsoft.Office.Interop.Excel;

namespace Office.Shared
{
	class ProtokolsOcher
	{
		protected string fileName;
		protected int oznakomLastString=49;
		protected Application app=new Application();
		protected bool visible;
		public ProtokolsOcher(string fileName, bool visible) {
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


			for (int rowIndex=1; rowIndex <= 41; rowIndex++) {
				string surname=wsPeople.Cells[1][rowIndex].Value.ToString();
				string firstName=wsPeople.Cells[2][rowIndex].Value.ToString();
				string secName=wsPeople.Cells[3][rowIndex].Value.ToString();

				string name=String.Format("{0} {1} {2}", surname, firstName, secName);
				string shortName=String.Format("{0} {1}. {2}.", surname, firstName.Substring(0, 1), secName.Substring(0, 1));
				string group=wsPeople.Cells[5][rowIndex].Value.ToString();
				string dolzn=wsPeople.Cells[7][rowIndex].Value.ToString();

				string blank = wsPeople.Cells[9][rowIndex].Value;
				Logger.log(name);
				wsBlank = wb.Worksheets[blank];

				wsBlank.Copy(newWS);
				Worksheet ws = newWB.Worksheets[blank];
				int len = shortName.Length < 30 ? shortName.Length : 30;
				ws.Name = String.Format("{0}", shortName.Substring(0, len));

				ws.Cells[3][13].Value = name;
				ws.Cells[3][15].Value = dolzn;
				ws.Cells[3][31].Value = String.Format("Оперативного персонала ({0})",dolzn);
				/*ws.Cells[3][18].Value = group;
				ws.Cells[3][29].Value = group;*/
				ws.Cells[9][44].Value = String.Format("/{0}/", shortName);

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
