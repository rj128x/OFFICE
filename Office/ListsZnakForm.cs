using System;
using System.Windows.Forms;
using Office.Shared;

namespace Office
{
	public partial class ListsZnakForm : Form
	{
		public ListsZnakForm() {
			InitializeComponent();
			Logger.addFunc(log);
		}
		private void log(string message) {
			txtLog.AppendText("\n" + message);
		}

		private void btnRun_Click(object sender, EventArgs e) {
			ListsZnak znak=new ListsZnak(txtFile.Text,chbVisible.Checked);
			znak.processFile();
		}

		private void btnChooseFile_Click(object sender, EventArgs e) {
			if (dlgFile.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFile.Text=dlgFile.FileName;
			}
		}

		private void btnCreateFolders_Click(object sender, EventArgs e) {
			InstrFolders folders=new InstrFolders(txtFile.Text, txtFolder.Text, chbVisible.Checked);
			folders.processFile();
		}

		private void btnChooseFolder_Click(object sender, EventArgs e) {
			if (dlgFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				txtFolder.Text = dlgFolder.SelectedPath;
			}
			
		}
	}
}
