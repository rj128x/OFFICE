using System;
using System.Windows.Forms;
using Office.Shared;

namespace Office
{
    public partial class OcencaForm : Form
	{
		public OcencaForm() {
			InitializeComponent();
			Logger.addFunc(log);
		}
		private void log(string message) {
			txtLog.AppendText("\n" + message);
		}

		private void btnRun_Click(object sender, EventArgs e) {
           Ocenca ocenca = new Ocenca(txtFile.Text, chbVisible.Checked);
		   ocenca.processFile();
		}

		private void btnChooseFile_Click(object sender, EventArgs e) {
			if (dlgFile.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFile.Text=dlgFile.FileName;
			}
		}        		
	}
}
