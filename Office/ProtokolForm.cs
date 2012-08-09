using System;
using System.Windows.Forms;
using Office.Shared;

namespace Office
{
    public partial class ProtokolForm : Form
	{
		public ProtokolForm() {
			InitializeComponent();
			Logger.addFunc(log);
		}
		private void log(string message) {
			txtLog.AppendText("\n" + message);
		}

		private void btnRun_Click(object sender, EventArgs e) {
           Protokols ocenca = new Protokols(txtFile.Text, chbVisible.Checked);
		   ocenca.processFile();
		}

		private void btnChooseFile_Click(object sender, EventArgs e) {
			if (dlgFile.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFile.Text=dlgFile.FileName;
			}
		}

		private void button1_Click(object sender, EventArgs e) {
			ProtokolsOcher ocenca = new ProtokolsOcher(txtFile.Text, chbVisible.Checked);
			ocenca.processFile();
		}        		
	}
}
