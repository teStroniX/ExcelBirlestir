namespace ExcelBirlestir
{
	partial class Form1
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
			this.label1 = new System.Windows.Forms.Label();
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
			this.label2 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
			this.SuspendLayout();
			// 
			// richTextBox1
			// 
			this.richTextBox1.Location = new System.Drawing.Point(12, 64);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(403, 140);
			this.richTextBox1.TabIndex = 0;
			this.richTextBox1.Text = "";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(421, 64);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 140);
			this.button1.TabIndex = 2;
			this.button1.Text = "Dosyaları Seç";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(12, 210);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(484, 23);
			this.button2.TabIndex = 3;
			this.button2.Text = "Çıktı Al";
			this.button2.UseVisualStyleBackColor = true;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// numericUpDown1
			// 
			this.numericUpDown1.Location = new System.Drawing.Point(395, 12);
			this.numericUpDown1.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.numericUpDown1.Name = "numericUpDown1";
			this.numericUpDown1.Size = new System.Drawing.Size(101, 20);
			this.numericUpDown1.TabIndex = 4;
			this.numericUpDown1.Value = new decimal(new int[] {
            24,
            0,
            0,
            0});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(309, 14);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(68, 13);
			this.label1.TabIndex = 5;
			this.label1.Text = "Sütun Sayısı:";
			// 
			// progressBar1
			// 
			this.progressBar1.Location = new System.Drawing.Point(12, 9);
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(291, 49);
			this.progressBar1.TabIndex = 6;
			// 
			// numericUpDown2
			// 
			this.numericUpDown2.Location = new System.Drawing.Point(395, 38);
			this.numericUpDown2.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.numericUpDown2.Name = "numericUpDown2";
			this.numericUpDown2.Size = new System.Drawing.Size(101, 20);
			this.numericUpDown2.TabIndex = 4;
			this.numericUpDown2.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(309, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 13);
			this.label2.TabIndex = 5;
			this.label2.Text = "Başlangıç Satır:";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(508, 245);
			this.Controls.Add(this.progressBar1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.numericUpDown2);
			this.Controls.Add(this.numericUpDown1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.richTextBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Form1";
			this.Text = "Excel Birleştir";
			this.TopMost = true;
			((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.RichTextBox richTextBox1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.NumericUpDown numericUpDown1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.NumericUpDown numericUpDown2;
		private System.Windows.Forms.Label label2;
	}
}

