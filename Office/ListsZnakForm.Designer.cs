namespace Office
{
	partial class ListsZnakForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.dlgFile = new System.Windows.Forms.OpenFileDialog();
			this.btnChooseFile = new System.Windows.Forms.Button();
			this.txtFile = new System.Windows.Forms.TextBox();
			this.btnRun = new System.Windows.Forms.Button();
			this.chbVisible = new System.Windows.Forms.CheckBox();
			this.txtLog = new System.Windows.Forms.RichTextBox();
			this.txtFolder = new System.Windows.Forms.TextBox();
			this.dlgFolder = new System.Windows.Forms.FolderBrowserDialog();
			this.btnChooseFolder = new System.Windows.Forms.Button();
			this.btnCreateFolders = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// dlgFile
			// 
			this.dlgFile.FileName = "openFileDialog1";
			// 
			// btnChooseFile
			// 
			this.btnChooseFile.Location = new System.Drawing.Point(430, 12);
			this.btnChooseFile.Name = "btnChooseFile";
			this.btnChooseFile.Size = new System.Drawing.Size(29, 23);
			this.btnChooseFile.TabIndex = 0;
			this.btnChooseFile.Text = "...";
			this.btnChooseFile.UseVisualStyleBackColor = true;
			this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
			// 
			// txtFile
			// 
			this.txtFile.Location = new System.Drawing.Point(20, 13);
			this.txtFile.Name = "txtFile";
			this.txtFile.Size = new System.Drawing.Size(404, 20);
			this.txtFile.TabIndex = 1;
			// 
			// btnRun
			// 
			this.btnRun.Location = new System.Drawing.Point(465, 10);
			this.btnRun.Name = "btnRun";
			this.btnRun.Size = new System.Drawing.Size(107, 23);
			this.btnRun.TabIndex = 2;
			this.btnRun.Text = "Создать листы";
			this.btnRun.UseVisualStyleBackColor = true;
			this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
			// 
			// chbVisible
			// 
			this.chbVisible.AutoSize = true;
			this.chbVisible.Location = new System.Drawing.Point(587, 14);
			this.chbVisible.Name = "chbVisible";
			this.chbVisible.Size = new System.Drawing.Size(52, 17);
			this.chbVisible.TabIndex = 3;
			this.chbVisible.Text = "Excel";
			this.chbVisible.UseVisualStyleBackColor = true;
			// 
			// txtLog
			// 
			this.txtLog.Location = new System.Drawing.Point(12, 105);
			this.txtLog.Name = "txtLog";
			this.txtLog.Size = new System.Drawing.Size(691, 207);
			this.txtLog.TabIndex = 4;
			this.txtLog.Text = "";
			// 
			// txtFolder
			// 
			this.txtFolder.Location = new System.Drawing.Point(20, 39);
			this.txtFolder.Name = "txtFolder";
			this.txtFolder.Size = new System.Drawing.Size(404, 20);
			this.txtFolder.TabIndex = 1;
			// 
			// btnChooseFolder
			// 
			this.btnChooseFolder.Location = new System.Drawing.Point(430, 36);
			this.btnChooseFolder.Name = "btnChooseFolder";
			this.btnChooseFolder.Size = new System.Drawing.Size(29, 23);
			this.btnChooseFolder.TabIndex = 0;
			this.btnChooseFolder.Text = "...";
			this.btnChooseFolder.UseVisualStyleBackColor = true;
			this.btnChooseFolder.Click += new System.EventHandler(this.btnChooseFolder_Click);
			// 
			// btnCreateFolders
			// 
			this.btnCreateFolders.Location = new System.Drawing.Point(465, 37);
			this.btnCreateFolders.Name = "btnCreateFolders";
			this.btnCreateFolders.Size = new System.Drawing.Size(107, 23);
			this.btnCreateFolders.TabIndex = 2;
			this.btnCreateFolders.Text = "Создать папки";
			this.btnCreateFolders.UseVisualStyleBackColor = true;
			this.btnCreateFolders.Click += new System.EventHandler(this.btnCreateFolders_Click);
			// 
			// ListsZnakForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(719, 324);
			this.Controls.Add(this.txtLog);
			this.Controls.Add(this.chbVisible);
			this.Controls.Add(this.btnCreateFolders);
			this.Controls.Add(this.btnRun);
			this.Controls.Add(this.txtFolder);
			this.Controls.Add(this.txtFile);
			this.Controls.Add(this.btnChooseFolder);
			this.Controls.Add(this.btnChooseFile);
			this.Name = "ListsZnakForm";
			this.Text = "ListsZnakForm";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.OpenFileDialog dlgFile;
		private System.Windows.Forms.Button btnChooseFile;
		private System.Windows.Forms.TextBox txtFile;
		private System.Windows.Forms.Button btnRun;
		private System.Windows.Forms.CheckBox chbVisible;
		private System.Windows.Forms.RichTextBox txtLog;
		private System.Windows.Forms.TextBox txtFolder;
		private System.Windows.Forms.FolderBrowserDialog dlgFolder;
		private System.Windows.Forms.Button btnChooseFolder;
		private System.Windows.Forms.Button btnCreateFolders;
	}
}