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
			maskedTxtEstabelecimento = new MaskedTextBox();
			btnImportar = new Button();
			label2 = new Label();
			openFileDialog1 = new OpenFileDialog();
			btnExcel = new Button();
			listView1 = new ListView();
			btnDelExcel = new Button();
			maskedTextBox2 = new MaskedTextBox();
			textBoxExcel1 = new TextBox();
			labelExcel1 = new Label();
			comboBoxImportacao = new ComboBox();
			label4 = new Label();
			labelExcel2 = new Label();
			textBoxExcel2 = new TextBox();
			btnExcel2 = new Button();
			label3 = new Label();
			comboBoxSistema = new ComboBox();
			SuspendLayout();
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Location = new Point(10, 44);
			label1.Name = "label1";
			label1.Size = new Size(108, 15);
			label1.TabIndex = 0;
			label1.Text = "EstabelecimentoID:";
			// 
			// maskedTxtEstabelecimento
			// 
			maskedTxtEstabelecimento.Location = new Point(136, 42);
			maskedTxtEstabelecimento.Margin = new Padding(3, 2, 3, 2);
			maskedTxtEstabelecimento.Mask = "00000000000";
			maskedTxtEstabelecimento.Name = "maskedTxtEstabelecimento";
			maskedTxtEstabelecimento.Size = new Size(110, 23);
			maskedTxtEstabelecimento.TabIndex = 2;
			maskedTxtEstabelecimento.ValidatingType = typeof(int);
			maskedTxtEstabelecimento.MaskInputRejected += maskedTxtEstabelecimento_MaskInputRejected;
			maskedTxtEstabelecimento.Click += maskedTxtEstabelecimento_Click;
			maskedTxtEstabelecimento.TextChanged += maskedTxtEstabelecimento_TextChanged;
			maskedTxtEstabelecimento.Enter += maskedTxtEstabelecimento_Enter;
			// 
			// btnImportar
			// 
			btnImportar.Location = new Point(306, 291);
			btnImportar.Margin = new Padding(3, 2, 3, 2);
			btnImportar.Name = "btnImportar";
			btnImportar.Size = new Size(88, 22);
			btnImportar.TabIndex = 5;
			btnImportar.Text = "⚙ Importar";
			btnImportar.UseVisualStyleBackColor = true;
			btnImportar.Click += btnImportar_Click;
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new Point(10, 194);
			label2.Name = "label2";
			label2.Size = new Size(74, 15);
			label2.TabIndex = 4;
			label2.Text = "ID Adicional:";
			label2.Visible = false;
			// 
			// openFileDialog1
			// 
			openFileDialog1.FileName = "openFileDialog1";
			// 
			// btnExcel
			// 
			btnExcel.Location = new Point(659, 68);
			btnExcel.Margin = new Padding(3, 2, 3, 2);
			btnExcel.Name = "btnExcel";
			btnExcel.Size = new Size(31, 22);
			btnExcel.TabIndex = 3;
			btnExcel.Text = "📂";
			btnExcel.UseVisualStyleBackColor = true;
			btnExcel.Visible = false;
			btnExcel.Click += btnExcel_Click;
			// 
			// listView1
			// 
			listView1.Location = new Point(551, 269);
			listView1.Margin = new Padding(3, 2, 3, 2);
			listView1.Name = "listView1";
			listView1.Size = new Size(139, 44);
			listView1.TabIndex = 3;
			listView1.UseCompatibleStateImageBehavior = false;
			listView1.View = View.List;
			listView1.Visible = false;
			// 
			// btnDelExcel
			// 
			btnDelExcel.Location = new Point(644, 235);
			btnDelExcel.Margin = new Padding(3, 2, 3, 2);
			btnDelExcel.Name = "btnDelExcel";
			btnDelExcel.Size = new Size(31, 22);
			btnDelExcel.TabIndex = 2;
			btnDelExcel.Text = "🗑";
			btnDelExcel.UseVisualStyleBackColor = true;
			btnDelExcel.Visible = false;
			btnDelExcel.Click += btnDelExcel_Click;
			// 
			// maskedTextBox2
			// 
			maskedTextBox2.Location = new Point(136, 194);
			maskedTextBox2.Margin = new Padding(3, 2, 3, 2);
			maskedTextBox2.Mask = "00000000000";
			maskedTextBox2.Name = "maskedTextBox2";
			maskedTextBox2.Size = new Size(110, 23);
			maskedTextBox2.TabIndex = 5;
			maskedTextBox2.ValidatingType = typeof(int);
			maskedTextBox2.Visible = false;
			// 
			// textBoxExcel1
			// 
			textBoxExcel1.Location = new Point(136, 67);
			textBoxExcel1.Margin = new Padding(3, 2, 3, 2);
			textBoxExcel1.Name = "textBoxExcel1";
			textBoxExcel1.Size = new Size(518, 23);
			textBoxExcel1.TabIndex = 7;
			textBoxExcel1.Visible = false;
			// 
			// labelExcel1
			// 
			labelExcel1.AutoSize = true;
			labelExcel1.Location = new Point(11, 69);
			labelExcel1.Name = "labelExcel1";
			labelExcel1.Size = new Size(108, 15);
			labelExcel1.TabIndex = 8;
			labelExcel1.Text = "EstabelecimentoID:";
			labelExcel1.Visible = false;
			// 
			// comboBoxImportacao
			// 
			comboBoxImportacao.DropDownStyle = ComboBoxStyle.DropDownList;
			comboBoxImportacao.FormattingEnabled = true;
			comboBoxImportacao.Items.AddRange(new object[] { "Pacientes" });
			comboBoxImportacao.Location = new Point(479, 9);
			comboBoxImportacao.Margin = new Padding(3, 2, 3, 2);
			comboBoxImportacao.Name = "comboBoxImportacao";
			comboBoxImportacao.Size = new Size(211, 23);
			comboBoxImportacao.TabIndex = 1;
			comboBoxImportacao.SelectedIndexChanged += comboBoxImportacao_SelectedIndexChanged;
			// 
			// label4
			// 
			label4.AutoSize = true;
			label4.Location = new Point(377, 11);
			label4.Name = "label4";
			label4.Size = new Size(87, 15);
			label4.TabIndex = 10;
			label4.Text = "Importação de:";
			// 
			// labelExcel2
			// 
			labelExcel2.AutoSize = true;
			labelExcel2.Location = new Point(11, 94);
			labelExcel2.Name = "labelExcel2";
			labelExcel2.Size = new Size(108, 15);
			labelExcel2.TabIndex = 11;
			labelExcel2.Text = "EstabelecimentoID:";
			labelExcel2.Visible = false;
			// 
			// textBoxExcel2
			// 
			textBoxExcel2.Location = new Point(136, 92);
			textBoxExcel2.Margin = new Padding(3, 2, 3, 2);
			textBoxExcel2.Name = "textBoxExcel2";
			textBoxExcel2.Size = new Size(518, 23);
			textBoxExcel2.TabIndex = 12;
			textBoxExcel2.Visible = false;
			// 
			// btnExcel2
			// 
			btnExcel2.Location = new Point(659, 93);
			btnExcel2.Margin = new Padding(3, 2, 3, 2);
			btnExcel2.Name = "btnExcel2";
			btnExcel2.Size = new Size(31, 22);
			btnExcel2.TabIndex = 4;
			btnExcel2.Text = "📂";
			btnExcel2.UseVisualStyleBackColor = true;
			btnExcel2.Visible = false;
			btnExcel2.Click += btnExcel2_Click;
			// 
			// label3
			// 
			label3.AutoSize = true;
			label3.Location = new Point(10, 11);
			label3.Name = "label3";
			label3.Size = new Size(89, 15);
			label3.TabIndex = 15;
			label3.Text = "Antigo sistema:";
			// 
			// comboBoxSistema
			// 
			comboBoxSistema.DropDownStyle = ComboBoxStyle.DropDownList;
			comboBoxSistema.FormattingEnabled = true;
			comboBoxSistema.Items.AddRange(new object[] { "DentalOffice", "OdontoCompany" });
			comboBoxSistema.Location = new Point(112, 9);
			comboBoxSistema.Margin = new Padding(3, 2, 3, 2);
			comboBoxSistema.Name = "comboBoxSistema";
			comboBoxSistema.Size = new Size(211, 23);
			comboBoxSistema.TabIndex = 0;
			comboBoxSistema.SelectedIndexChanged += comboBoxSistema_SelectedIndexChanged;
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(700, 322);
			Controls.Add(label3);
			Controls.Add(comboBoxSistema);
			Controls.Add(btnExcel2);
			Controls.Add(textBoxExcel2);
			Controls.Add(labelExcel2);
			Controls.Add(label4);
			Controls.Add(comboBoxImportacao);
			Controls.Add(labelExcel1);
			Controls.Add(textBoxExcel1);
			Controls.Add(maskedTextBox2);
			Controls.Add(btnDelExcel);
			Controls.Add(listView1);
			Controls.Add(btnExcel);
			Controls.Add(label2);
			Controls.Add(btnImportar);
			Controls.Add(maskedTxtEstabelecimento);
			Controls.Add(label1);
			FormBorderStyle = FormBorderStyle.FixedToolWindow;
			Margin = new Padding(3, 2, 3, 2);
			Name = "Form1";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "Migração";
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label label1;
		private MaskedTextBox maskedTxtEstabelecimento;
		private Button btnImportar;
		private Label label2;
		private OpenFileDialog openFileDialog1;
		private Button btnExcel;
		private ListView listView1;
		private Button btnDelExcel;
		private MaskedTextBox maskedTextBox2;
		private TextBox textBoxExcel1;
		private Label labelExcel1;
		private ComboBox comboBoxImportacao;
		private Label label4;
		private Label labelExcel2;
		private TextBox textBoxExcel2;
		private Button btnExcel2;
		private Label label3;
		private ComboBox comboBoxSistema;
	}
}
