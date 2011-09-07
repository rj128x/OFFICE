namespace Office
{
	partial class BlanksForm
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
			this.btnBlankCurrent = new System.Windows.Forms.Button();
			this.chbShowWord = new System.Windows.Forms.CheckBox();
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
			this.richTextBox1.Location = new System.Drawing.Point(12, 80);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(458, 259);
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
			this.btnBlankTip.Location = new System.Drawing.Point(13, 39);
			this.btnBlankTip.Name = "btnBlankTip";
			this.btnBlankTip.Size = new System.Drawing.Size(105, 23);
			this.btnBlankTip.TabIndex = 3;
			this.btnBlankTip.Text = "Типовые бланки";
			this.btnBlankTip.UseVisualStyleBackColor = true;
			this.btnBlankTip.Click += new System.EventHandler(this.btnBlankTip_Click);
			// 
			// btnBlankCurrent
			// 
			this.btnBlankCurrent.Location = new System.Drawing.Point(124, 39);
			this.btnBlankCurrent.Name = "btnBlankCurrent";
			this.btnBlankCurrent.Size = new System.Drawing.Size(105, 23);
			this.btnBlankCurrent.TabIndex = 3;
			this.btnBlankCurrent.Text = "Текущие бланки";
			this.btnBlankCurrent.UseVisualStyleBackColor = true;
			this.btnBlankCurrent.Click += new System.EventHandler(this.btnBlankCurrent_Click);
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
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(510, 351);
			this.Controls.Add(this.chbShowWord);
			this.Controls.Add(this.btnBlankCurrent);
			this.Controls.Add(this.btnBlankTip);
			this.Controls.Add(this.txtFolder);
			this.Controls.Add(this.richTextBox1);
			this.Controls.Add(this.btnSelectFolder);
			this.Name = "Form1";
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
		private System.Windows.Forms.Button btnBlankCurrent;
		private System.Windows.Forms.CheckBox chbShowWord;
	}
}

