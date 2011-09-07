using System;
using System.Windows.Forms;
using Office.Shared;

namespace Office
{
	public partial class BlanksForm : Form
	{
		public BlanksForm() {
			InitializeComponent();
			Logger.addFunc(log);
		}
		private void log(string message) {
			richTextBox1.AppendText("\n" + message);
		}

		private void btnSelectFolder_Click(object sender, EventArgs e) {
			if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolder.Text=folderBrowserDialog.SelectedPath;
			}
		}

		private void btnBlankTip_Click(object sender, EventArgs e) {
			ProcessBlanks blanks=new ProcessBlanks(txtFolder.Text, BlankOperation.tip,chbShowWord.Checked);
		}

		private void btnBlankCurrent_Click(object sender, EventArgs e) {
			ProcessBlanks blanks=new ProcessBlanks(txtFolder.Text, BlankOperation.current, chbShowWord.Checked);
		}
	}
}
