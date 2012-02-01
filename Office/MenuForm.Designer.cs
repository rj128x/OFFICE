namespace Office
{
	partial class MenuForm
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
			this.btnBlanks = new System.Windows.Forms.Button();
			this.btnListsZnak = new System.Windows.Forms.Button();
			this.btnOcenca = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnBlanks
			// 
			this.btnBlanks.Location = new System.Drawing.Point(12, 12);
			this.btnBlanks.Name = "btnBlanks";
			this.btnBlanks.Size = new System.Drawing.Size(75, 23);
			this.btnBlanks.TabIndex = 0;
			this.btnBlanks.Text = "Бланки";
			this.btnBlanks.UseVisualStyleBackColor = true;
			this.btnBlanks.Click += new System.EventHandler(this.btnBlanks_Click);
			// 
			// btnListsZnak
			// 
			this.btnListsZnak.Location = new System.Drawing.Point(12, 41);
			this.btnListsZnak.Name = "btnListsZnak";
			this.btnListsZnak.Size = new System.Drawing.Size(75, 23);
			this.btnListsZnak.TabIndex = 0;
			this.btnListsZnak.Text = "Листы";
			this.btnListsZnak.UseVisualStyleBackColor = true;
			this.btnListsZnak.Click += new System.EventHandler(this.btnListsZnak_Click);
			// 
			// btnOcenca
			// 
			this.btnOcenca.Location = new System.Drawing.Point(12, 70);
			this.btnOcenca.Name = "btnOcenca";
			this.btnOcenca.Size = new System.Drawing.Size(75, 23);
			this.btnOcenca.TabIndex = 0;
			this.btnOcenca.Text = "протоколы";
			this.btnOcenca.UseVisualStyleBackColor = true;
			this.btnOcenca.Click += new System.EventHandler(this.btnOcenca_Click);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(12, 99);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 23);
			this.button1.TabIndex = 0;
			this.button1.Text = "пдф";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// MenuForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(284, 262);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnOcenca);
			this.Controls.Add(this.btnListsZnak);
			this.Controls.Add(this.btnBlanks);
			this.Name = "MenuForm";
			this.Text = "MenuForm";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnBlanks;
		private System.Windows.Forms.Button btnListsZnak;
        private System.Windows.Forms.Button btnOcenca;
		  private System.Windows.Forms.Button button1;
	}
}