using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office.Shared;

namespace Office
{
	public partial class CreatePDFForm : Form
	{
		public CreatePDFForm() {
			InitializeComponent();
		}
		private void log(string message) {
			richTextBox1.AppendText("\n" + message);
		}

		private void btnSelectFolder_Click(object sender, EventArgs e) {
			if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolder.Text = folderBrowserDialog.SelectedPath;
			}
		}

		private void btnBlankTip_Click(object sender, EventArgs e) {
			ProcessPDF pdf=new ProcessPDF(txtFolder.Text, txtFolderPDF.Text,chb1Page.Checked,  chbShowWord.Checked);
		}



		private void btnSelectFolderPDF_Click(object sender, EventArgs e) {
			if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolderPDF.Text = folderBrowserDialog.SelectedPath;
			}
		}

		
	}
}
