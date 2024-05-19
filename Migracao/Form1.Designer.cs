namespace Migracao
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			label1 = new Label();
			btnImportar = new Button();
			label2 = new Label();
			openFileDialog1 = new OpenFileDialog();
			btnExcel = new Button();
			textBoxExcel1 = new TextBox();
			labelExcel1 = new Label();
			comboBoxImportacao = new ComboBox();
			label4 = new Label();
			labelExcel2 = new Label();
			textBoxExcel2 = new TextBox();
			btnExcel2 = new Button();
			label3 = new Label();
			comboBoxSistema = new ComboBox();
			txtEstabelecimentoID = new TextBox();
			txtPessoaID = new TextBox();
			txtLoginID = new TextBox();
			label5 = new Label();
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
			// btnImportar
			// 
			btnImportar.Location = new Point(293, 204);
			btnImportar.Margin = new Padding(3, 2, 3, 2);
			btnImportar.Name = "btnImportar";
			btnImportar.Size = new Size(88, 22);
			btnImportar.TabIndex = 6;
			btnImportar.Text = "⚙ Importar";
			btnImportar.UseVisualStyleBackColor = true;
			btnImportar.Click += btnImportar_Click;
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new Point(11, 149);
			label2.Name = "label2";
			label2.Size = new Size(104, 15);
			label2.TabIndex = 4;
			label2.Text = "PessoaID Resp Fin:";
			label2.Visible = false;
			// 
			// openFileDialog1
			// 
			openFileDialog1.FileName = "openFileDialog1";
			// 
			// btnExcel
			// 
			btnExcel.Location = new Point(633, 97);
			btnExcel.Margin = new Padding(3, 2, 3, 2);
			btnExcel.Name = "btnExcel";
			btnExcel.Size = new Size(31, 22);
			btnExcel.TabIndex = 3;
			btnExcel.Text = "📂";
			btnExcel.UseVisualStyleBackColor = true;
			btnExcel.Visible = false;
			btnExcel.Click += btnExcel_Click;
			// 
			// textBoxExcel1
			// 
			textBoxExcel1.Location = new Point(132, 96);
			textBoxExcel1.Margin = new Padding(3, 2, 3, 2);
			textBoxExcel1.Name = "textBoxExcel1";
			textBoxExcel1.Size = new Size(495, 23);
			textBoxExcel1.TabIndex = 7;
			textBoxExcel1.Visible = false;
			// 
			// labelExcel1
			// 
			labelExcel1.AutoSize = true;
			labelExcel1.Location = new Point(11, 99);
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
			comboBoxImportacao.Items.AddRange(new object[] { "Pacientes", "Recebidos", "Fornecedores" });
			comboBoxImportacao.Location = new Point(453, 13);
			comboBoxImportacao.Margin = new Padding(3, 2, 3, 2);
			comboBoxImportacao.Name = "comboBoxImportacao";
			comboBoxImportacao.Size = new Size(211, 23);
			comboBoxImportacao.TabIndex = 1;
			comboBoxImportacao.SelectedIndexChanged += comboBoxImportacao_SelectedIndexChanged;
			// 
			// label4
			// 
			label4.AutoSize = true;
			label4.Location = new Point(360, 16);
			label4.Name = "label4";
			label4.Size = new Size(87, 15);
			label4.TabIndex = 10;
			label4.Text = "Importação de:";
			// 
			// labelExcel2
			// 
			labelExcel2.AutoSize = true;
			labelExcel2.Location = new Point(11, 126);
			labelExcel2.Name = "labelExcel2";
			labelExcel2.Size = new Size(108, 15);
			labelExcel2.TabIndex = 11;
			labelExcel2.Text = "EstabelecimentoID:";
			labelExcel2.Visible = false;
			// 
			// textBoxExcel2
			// 
			textBoxExcel2.Location = new Point(132, 123);
			textBoxExcel2.Margin = new Padding(3, 2, 3, 2);
			textBoxExcel2.Name = "textBoxExcel2";
			textBoxExcel2.Size = new Size(495, 23);
			textBoxExcel2.TabIndex = 12;
			textBoxExcel2.Visible = false;
			// 
			// btnExcel2
			// 
			btnExcel2.Location = new Point(633, 124);
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
			label3.Location = new Point(12, 16);
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
			comboBoxSistema.Location = new Point(132, 11);
			comboBoxSistema.Margin = new Padding(3, 2, 3, 2);
			comboBoxSistema.Name = "comboBoxSistema";
			comboBoxSistema.Size = new Size(211, 23);
			comboBoxSistema.TabIndex = 0;
			comboBoxSistema.SelectedIndexChanged += comboBoxSistema_SelectedIndexChanged;
			// 
			// txtEstabelecimentoID
			// 
			txtEstabelecimentoID.Location = new Point(132, 41);
			txtEstabelecimentoID.Margin = new Padding(3, 2, 3, 2);
			txtEstabelecimentoID.Name = "txtEstabelecimentoID";
			txtEstabelecimentoID.Size = new Size(110, 23);
			txtEstabelecimentoID.TabIndex = 2;
			txtEstabelecimentoID.TextChanged += txtEstabelecimentoID_TextChanged;
			txtEstabelecimentoID.KeyPress += txtEstabelecimentoID_KeyPress;
			// 
			// txtPessoaID
			// 
			txtPessoaID.Location = new Point(132, 150);
			txtPessoaID.Margin = new Padding(3, 2, 3, 2);
			txtPessoaID.Name = "txtPessoaID";
			txtPessoaID.Size = new Size(110, 23);
			txtPessoaID.TabIndex = 5;
			txtPessoaID.Visible = false;
			txtPessoaID.TextChanged += txtPessoaID_TextChanged;
			txtPessoaID.KeyPress += txtPessoaID_KeyPress;
			// 
			// txtLoginID
			// 
			txtLoginID.Location = new Point(132, 69);
			txtLoginID.Margin = new Padding(3, 2, 3, 2);
			txtLoginID.Name = "txtLoginID";
			txtLoginID.Size = new Size(110, 23);
			txtLoginID.TabIndex = 17;
			txtLoginID.Text = "1";
			txtLoginID.KeyPress += txtLoginID_KeyPress;
			// 
			// label5
			// 
			label5.AutoSize = true;
			label5.Location = new Point(12, 72);
			label5.Name = "label5";
			label5.Size = new Size(51, 15);
			label5.TabIndex = 16;
			label5.Text = "LoginID:";
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(675, 237);
			Controls.Add(txtLoginID);
			Controls.Add(label5);
			Controls.Add(txtPessoaID);
			Controls.Add(txtEstabelecimentoID);
			Controls.Add(label3);
			Controls.Add(comboBoxSistema);
			Controls.Add(btnExcel2);
			Controls.Add(textBoxExcel2);
			Controls.Add(labelExcel2);
			Controls.Add(label4);
			Controls.Add(comboBoxImportacao);
			Controls.Add(labelExcel1);
			Controls.Add(textBoxExcel1);
			Controls.Add(btnExcel);
			Controls.Add(label2);
			Controls.Add(btnImportar);
			Controls.Add(label1);
			FormBorderStyle = FormBorderStyle.FixedToolWindow;
			Icon = (Icon)resources.GetObject("$this.Icon");
			Margin = new Padding(3, 2, 3, 2);
			Name = "Form1";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "Migração";
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label label1;
		private Button btnImportar;
		private Label label2;
		private OpenFileDialog openFileDialog1;
		private Button btnExcel;
		private TextBox textBoxExcel1;
		private Label labelExcel1;
		private ComboBox comboBoxImportacao;
		private Label label4;
		private Label labelExcel2;
		private TextBox textBoxExcel2;
		private Button btnExcel2;
		private Label label3;
		private ComboBox comboBoxSistema;
		private TextBox txtEstabelecimentoID;
		private TextBox txtPessoaID;
		private TextBox txtLoginID;
		private Label label5;
	}
}
