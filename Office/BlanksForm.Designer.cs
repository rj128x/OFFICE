namespace Office
{
	partial class PDFForm
	{
		/// <summary>
		/// Требуется переменная конструктора.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		/// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Код, автоматически созданный конструктором форм Windows

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое данного метода при помощи редактора кода.
		/// </summary>
		private void InitializeComponent() {
			this.btnSelectFolder = new System.Windows.Forms.Button();
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
			this.txtFolder = new System.Windows.Forms.TextBox();
			this.btnBlankTip = new System.Windows.Forms.Button();
			this.chbShowWord = new System.Windows.Forms.CheckBox();
			this.chbCreatePDF = new System.Windows.Forms.CheckBox();
			this.txtFolderPDF = new System.Windows.Forms.TextBox();
			this.btnSelectFolderPDF = new System.Windows.Forms.Button();
			this.txtFolderTip = new System.Windows.Forms.TextBox();
			this.btnSelectFolderTip = new System.Windows.Forms.Button();
			this.chbCreateTip = new System.Windows.Forms.CheckBox();
			this.btnSelectFolderCurrent = new System.Windows.Forms.Button();
			this.txtFolderCurrent = new System.Windows.Forms.TextBox();
			this.chbCreateCurrent = new System.Windows.Forms.CheckBox();
			this.chb1Page = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// btnSelectFolder
			// 
			this.btnSelectFolder.Location = new System.Drawing.Point(387, 10);
			this.btnSelectFolder.Name = "btnSelectFolder";
			this.btnSelectFolder.Size = new System.Drawing.Size(26, 23);
			this.btnSelectFolder.TabIndex = 0;
			this.btnSelectFolder.Text = "...";
			this.btnSelectFolder.UseVisualStyleBackColor = true;
			this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
			// 
			// richTextBox1
			// 
			this.richTextBox1.Location = new System.Drawing.Point(12, 143);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(458, 196);
			this.richTextBox1.TabIndex = 1;
			this.richTextBox1.Text = "";
			// 
			// txtFolder
			// 
			this.txtFolder.Location = new System.Drawing.Point(12, 12);
			this.txtFolder.Name = "txtFolder";
			this.txtFolder.Size = new System.Drawing.Size(369, 20);
			this.txtFolder.TabIndex = 2;
			// 
			// btnBlankTip
			// 
			this.btnBlankTip.Location = new System.Drawing.Point(12, 114);
			this.btnBlankTip.Name = "btnBlankTip";
			this.btnBlankTip.Size = new System.Drawing.Size(105, 23);
			this.btnBlankTip.TabIndex = 3;
			this.btnBlankTip.Text = "Типовые бланки";
			this.btnBlankTip.UseVisualStyleBackColor = true;
			this.btnBlankTip.Click += new System.EventHandler(this.btnBlankTip_Click);
			// 
			// chbShowWord
			// 
			this.chbShowWord.AutoSize = true;
			this.chbShowWord.Location = new System.Drawing.Point(420, 15);
			this.chbShowWord.Name = "chbShowWord";
			this.chbShowWord.Size = new System.Drawing.Size(52, 17);
			this.chbShowWord.TabIndex = 4;
			this.chbShowWord.Text = "Word";
			this.chbShowWord.UseVisualStyleBackColor = true;
			// 
			// chbCreatePDF
			// 
			this.chbCreatePDF.AutoSize = true;
			this.chbCreatePDF.Location = new System.Drawing.Point(420, 38);
			this.chbCreatePDF.Name = "chbCreatePDF";
			this.chbCreatePDF.Size = new System.Drawing.Size(41, 17);
			this.chbCreatePDF.TabIndex = 4;
			this.chbCreatePDF.Text = "pdf";
			this.chbCreatePDF.UseVisualStyleBackColor = true;
			// 
			// txtFolderPDF
			// 
			this.txtFolderPDF.Location = new System.Drawing.Point(12, 36);
			this.txtFolderPDF.Name = "txtFolderPDF";
			this.txtFolderPDF.Size = new System.Drawing.Size(369, 20);
			this.txtFolderPDF.TabIndex = 2;
			// 
			// btnSelectFolderPDF
			// 
			this.btnSelectFolderPDF.Location = new System.Drawing.Point(387, 34);
			this.btnSelectFolderPDF.Name = "btnSelectFolderPDF";
			this.btnSelectFolderPDF.Size = new System.Drawing.Size(26, 23);
			this.btnSelectFolderPDF.TabIndex = 0;
			this.btnSelectFolderPDF.Text = "...";
			this.btnSelectFolderPDF.UseVisualStyleBackColor = true;
			this.btnSelectFolderPDF.Click += new System.EventHandler(this.btnSelectFolderPDF_Click);
			// 
			// txtFolderTip
			// 
			this.txtFolderTip.Location = new System.Drawing.Point(12, 62);
			this.txtFolderTip.Name = "txtFolderTip";
			this.txtFolderTip.Size = new System.Drawing.Size(369, 20);
			this.txtFolderTip.TabIndex = 2;
			// 
			// btnSelectFolderTip
			// 
			this.btnSelectFolderTip.Location = new System.Drawing.Point(387, 59);
			this.btnSelectFolderTip.Name = "btnSelectFolderTip";
			this.btnSelectFolderTip.Size = new System.Drawing.Size(26, 23);
			this.btnSelectFolderTip.TabIndex = 0;
			this.btnSelectFolderTip.Text = "...";
			this.btnSelectFolderTip.UseVisualStyleBackColor = true;
			this.btnSelectFolderTip.Click += new System.EventHandler(this.btnSelectFolderTip_Click);
			// 
			// chbCreateTip
			// 
			this.chbCreateTip.AutoSize = true;
			this.chbCreateTip.Location = new System.Drawing.Point(420, 63);
			this.chbCreateTip.Name = "chbCreateTip";
			this.chbCreateTip.Size = new System.Drawing.Size(41, 17);
			this.chbCreateTip.TabIndex = 4;
			this.chbCreateTip.Text = "Tip";
			this.chbCreateTip.UseVisualStyleBackColor = true;
			// 
			// btnSelectFolderCurrent
			// 
			this.btnSelectFolderCurrent.Location = new System.Drawing.Point(387, 85);
			this.btnSelectFolderCurrent.Name = "btnSelectFolderCurrent";
			this.btnSelectFolderCurrent.Size = new System.Drawing.Size(26, 23);
			this.btnSelectFolderCurrent.TabIndex = 0;
			this.btnSelectFolderCurrent.Text = "...";
			this.btnSelectFolderCurrent.UseVisualStyleBackColor = true;
			this.btnSelectFolderCurrent.Click += new System.EventHandler(this.btnSelectFolderCurrent_Click);
			// 
			// txtFolderCurrent
			// 
			this.txtFolderCurrent.Location = new System.Drawing.Point(12, 88);
			this.txtFolderCurrent.Name = "txtFolderCurrent";
			this.txtFolderCurrent.Size = new System.Drawing.Size(369, 20);
			this.txtFolderCurrent.TabIndex = 2;
			// 
			// chbCreateCurrent
			// 
			this.chbCreateCurrent.AutoSize = true;
			this.chbCreateCurrent.Location = new System.Drawing.Point(420, 89);
			this.chbCreateCurrent.Name = "chbCreateCurrent";
			this.chbCreateCurrent.Size = new System.Drawing.Size(60, 17);
			this.chbCreateCurrent.TabIndex = 4;
			this.chbCreateCurrent.Text = "Current";
			this.chbCreateCurrent.UseVisualStyleBackColor = true;
			// 
			// chb1Page
			// 
			this.chb1Page.AutoSize = true;
			this.chb1Page.Location = new System.Drawing.Point(457, 38);
			this.chb1Page.Name = "chb1Page";
			this.chb1Page.Size = new System.Drawing.Size(59, 17);
			this.chb1Page.TabIndex = 4;
			this.chb1Page.Text = "1 page";
			this.chb1Page.UseVisualStyleBackColor = true;
			// 
			// PDFForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(510, 351);
			this.Controls.Add(this.chbCreateCurrent);
			this.Controls.Add(this.chbCreateTip);
			this.Controls.Add(this.chb1Page);
			this.Controls.Add(this.chbCreatePDF);
			this.Controls.Add(this.chbShowWord);
			this.Controls.Add(this.btnBlankTip);
			this.Controls.Add(this.txtFolderCurrent);
			this.Controls.Add(this.txtFolderTip);
			this.Controls.Add(this.txtFolderPDF);
			this.Controls.Add(this.txtFolder);
			this.Controls.Add(this.btnSelectFolderCurrent);
			this.Controls.Add(this.richTextBox1);
			this.Controls.Add(this.btnSelectFolderTip);
			this.Controls.Add(this.btnSelectFolderPDF);
			this.Controls.Add(this.btnSelectFolder);
			this.Name = "PDFForm";
			this.Text = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnSelectFolder;
		private System.Windows.Forms.RichTextBox richTextBox1;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
		private System.Windows.Forms.TextBox txtFolder;
		private System.Windows.Forms.Button btnBlankTip;
		private System.Windows.Forms.CheckBox chbShowWord;
		private System.Windows.Forms.CheckBox chbCreatePDF;
		private System.Windows.Forms.TextBox txtFolderPDF;
		private System.Windows.Forms.Button btnSelectFolderPDF;
		private System.Windows.Forms.TextBox txtFolderTip;
		private System.Windows.Forms.Button btnSelectFolderTip;
		private System.Windows.Forms.CheckBox chbCreateTip;
		private System.Windows.Forms.Button btnSelectFolderCurrent;
		private System.Windows.Forms.TextBox txtFolderCurrent;
		private System.Windows.Forms.CheckBox chbCreateCurrent;
		private System.Windows.Forms.CheckBox chb1Page;
	}
}

