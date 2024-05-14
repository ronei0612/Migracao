namespace Migração
{
	partial class Form1
	{
		/// <summary>
		///  Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		///  Clean up any resources being used.
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
		///  Required method for Designer support - do not modify
		///  the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			label1 = new Label();
			maskedTextBox1 = new MaskedTextBox();
			btnImportar = new Button();
			label2 = new Label();
			openFileDialog1 = new OpenFileDialog();
			btnExcel = new Button();
			listView1 = new ListView();
			btnDelExcel = new Button();
			maskedTextBox2 = new MaskedTextBox();
			SuspendLayout();
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Location = new Point(13, 179);
			label1.Name = "label1";
			label1.Size = new Size(137, 20);
			label1.TabIndex = 0;
			label1.Text = "EstabelecimentoID:";
			// 
			// maskedTextBox1
			// 
			maskedTextBox1.Location = new Point(156, 176);
			maskedTextBox1.Mask = "00000000000";
			maskedTextBox1.Name = "maskedTextBox1";
			maskedTextBox1.Size = new Size(125, 27);
			maskedTextBox1.TabIndex = 4;
			maskedTextBox1.ValidatingType = typeof(int);
			// 
			// btnImportar
			// 
			btnImportar.Location = new Point(350, 244);
			btnImportar.Name = "btnImportar";
			btnImportar.Size = new Size(101, 29);
			btnImportar.TabIndex = 6;
			btnImportar.Text = "⚙ Importar";
			btnImportar.UseVisualStyleBackColor = true;
			btnImportar.Click += btnImportar_Click;
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new Point(13, 209);
			label2.Name = "label2";
			label2.Size = new Size(94, 20);
			label2.TabIndex = 4;
			label2.Text = "ID Adicional:";
			// 
			// openFileDialog1
			// 
			openFileDialog1.FileName = "openFileDialog1";
			// 
			// btnExcel
			// 
			btnExcel.Location = new Point(12, 12);
			btnExcel.Name = "btnExcel";
			btnExcel.Size = new Size(35, 29);
			btnExcel.TabIndex = 1;
			btnExcel.Text = "➕";
			btnExcel.UseVisualStyleBackColor = true;
			btnExcel.Click += btnExcel_Click;
			// 
			// listView1
			// 
			listView1.Location = new Point(12, 47);
			listView1.Name = "listView1";
			listView1.Size = new Size(776, 119);
			listView1.TabIndex = 3;
			listView1.UseCompatibleStateImageBehavior = false;
			listView1.View = View.List;
			// 
			// btnDelExcel
			// 
			btnDelExcel.Location = new Point(53, 12);
			btnDelExcel.Name = "btnDelExcel";
			btnDelExcel.Size = new Size(35, 29);
			btnDelExcel.TabIndex = 2;
			btnDelExcel.Text = "🗑";
			btnDelExcel.UseVisualStyleBackColor = true;
			btnDelExcel.Click += btnDelExcel_Click;
			// 
			// maskedTextBox2
			// 
			maskedTextBox2.Location = new Point(156, 209);
			maskedTextBox2.Mask = "00000000000";
			maskedTextBox2.Name = "maskedTextBox2";
			maskedTextBox2.Size = new Size(125, 27);
			maskedTextBox2.TabIndex = 5;
			maskedTextBox2.ValidatingType = typeof(int);
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(8F, 20F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(800, 288);
			Controls.Add(maskedTextBox2);
			Controls.Add(btnDelExcel);
			Controls.Add(listView1);
			Controls.Add(btnExcel);
			Controls.Add(label2);
			Controls.Add(btnImportar);
			Controls.Add(maskedTextBox1);
			Controls.Add(label1);
			FormBorderStyle = FormBorderStyle.FixedToolWindow;
			Name = "Form1";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "Migração";
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label label1;
		private MaskedTextBox maskedTextBox1;
		private Button btnImportar;
		private Label label2;
		private OpenFileDialog openFileDialog1;
		private Button btnExcel;
		private ListView listView1;
		private Button btnDelExcel;
		private MaskedTextBox maskedTextBox2;
	}
}
