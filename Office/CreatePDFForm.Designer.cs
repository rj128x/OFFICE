namespace Office
{
	partial class CreatePDFForm
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
			this.chbShowWord = new System.Windows.Forms.CheckBox();
			this.btnBlankTip = new System.Windows.Forms.Button();
			this.txtFolderPDF = new System.Windows.Forms.TextBox();
			this.txtFolder = new System.Windows.Forms.TextBox();
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.btnSelectFolderPDF = new System.Windows.Forms.Button();
			this.btnSelectFolder = new System.Windows.Forms.Button();
			this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
			this.chb1Page = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// chbShowWord
			// 
			this.chbShowWord.AutoSize = true;
			this.chbShowWord.Location = new System.Drawing.Point(420, 15);
			this.chbShowWord.Name = "chbShowWord";
			this.chbShowWord.Size = new System.Drawing.Size(52, 17);
			this.chbShowWord.TabIndex = 11;
			this.chbShowWord.Text = "Word";
			this.chbShowWord.UseVisualStyleBackColor = true;
			// 
			// btnBlankTip
			// 
			this.btnBlankTip.Location = new System.Drawing.Point(12, 62);
			this.btnBlankTip.Name = "btnBlankTip";
			this.btnBlankTip.Size = new System.Drawing.Size(105, 23);
			this.btnBlankTip.TabIndex = 10;
			this.btnBlankTip.Text = "Типовые бланки";
			this.btnBlankTip.UseVisualStyleBackColor = true;
			this.btnBlankTip.Click += new System.EventHandler(this.btnBlankTip_Click);
			// 
			// txtFolderPDF
			// 
			this.txtFolderPDF.Location = new System.Drawing.Point(12, 36);
			this.txtFolderPDF.Name = "txtFolderPDF";
			this.txtFolderPDF.Size = new System.Drawing.Size(369, 20);
			this.txtFolderPDF.TabIndex = 9;
			// 
			// txtFolder
			// 
			this.txtFolder.Location = new System.Drawing.Point(12, 12);
			this.txtFolder.Name = "txtFolder";
			this.txtFolder.Size = new System.Drawing.Size(369, 20);
			this.txtFolder.TabIndex = 8;
			// 
			// richTextBox1
			// 
			this.richTextBox1.Location = new System.Drawing.Point(12, 91);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(458, 196);
			this.richTextBox1.TabIndex = 7;
			this.richTextBox1.Text = "";
			// 
			// btnSelectFolderPDF
			// 
			this.btnSelectFolderPDF.Location = new System.Drawing.Point(387, 34);
			this.btnSelectFolderPDF.Name = "btnSelectFolderPDF";
			this.btnSelectFolderPDF.Size = new System.Drawing.Size(26, 23);
			this.btnSelectFolderPDF.TabIndex = 5;
			this.btnSelectFolderPDF.Text = "...";
			this.btnSelectFolderPDF.UseVisualStyleBackColor = true;
			this.btnSelectFolderPDF.Click += new System.EventHandler(this.btnSelectFolderPDF_Click);
			// 
			// btnSelectFolder
			// 
			this.btnSelectFolder.Location = new System.Drawing.Point(387, 10);
			this.btnSelectFolder.Name = "btnSelectFolder";
			this.btnSelectFolder.Size = new System.Drawing.Size(26, 23);
			this.btnSelectFolder.TabIndex = 6;
			this.btnSelectFolder.Text = "...";
			this.btnSelectFolder.UseVisualStyleBackColor = true;
			this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
			// 
			// chb1Page
			// 
			this.chb1Page.AutoSize = true;
			this.chb1Page.Location = new System.Drawing.Point(420, 38);
			this.chb1Page.Name = "chb1Page";
			this.chb1Page.Size = new System.Drawing.Size(59, 17);
			this.chb1Page.TabIndex = 12;
			this.chb1Page.Text = "1 page";
			this.chb1Page.UseVisualStyleBackColor = true;
			// 
			// CreatePDFForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(504, 314);
			this.Controls.Add(this.chb1Page);
			this.Controls.Add(this.chbShowWord);
			this.Controls.Add(this.btnBlankTip);
			this.Controls.Add(this.txtFolderPDF);
			this.Controls.Add(this.txtFolder);
			this.Controls.Add(this.richTextBox1);
			this.Controls.Add(this.btnSelectFolderPDF);
			this.Controls.Add(this.btnSelectFolder);
			this.Name = "CreatePDFForm";
			this.Text = "CreatePDFForm";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.CheckBox chbShowWord;
		private System.Windows.Forms.Button btnBlankTip;
		private System.Windows.Forms.TextBox txtFolderPDF;
		private System.Windows.Forms.TextBox txtFolder;
		private System.Windows.Forms.RichTextBox richTextBox1;
		private System.Windows.Forms.Button btnSelectFolderPDF;
		private System.Windows.Forms.Button btnSelectFolder;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
		private System.Windows.Forms.CheckBox chb1Page;
	}
}