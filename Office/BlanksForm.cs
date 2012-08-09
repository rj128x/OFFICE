using System;
using System.Windows.Forms;
using Office.Shared;

namespace Office
{
	public partial class PDFForm : Form
	{
		public PDFForm() {
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
			ProcessBlanks blanks=new ProcessBlanks(txtFolder.Text,txtFolderPDF.Text,txtFolderTip.Text,txtFolderCurrent.Text,
				chbCreatePDF.Checked,chb1Page.Checked,chbCreateTip.Checked,chbCreateCurrent.Checked,chbShowWord.Checked);
		}



		private void btnSelectFolderPDF_Click(object sender, EventArgs e) {
			if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolderPDF.Text = folderBrowserDialog.SelectedPath;
			}
		}

		private void btnSelectFolderTip_Click(object sender, EventArgs e) {
			if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolderTip.Text = folderBrowserDialog.SelectedPath;
			}
		}

		private void btnSelectFolderCurrent_Click(object sender, EventArgs e) {
			if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolderCurrent.Text = folderBrowserDialog.SelectedPath;
			}
		}
	}
}
