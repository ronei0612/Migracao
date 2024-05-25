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
			lbPessoaID = new Label();
			openFileDialog1 = new OpenFileDialog();
			btnExcel = new Button();
			textBoxExcel1 = new TextBox();
			labelExcel1 = new Label();
			comboBoxImportacao = new ComboBox();
			label4 = new Label();
			lbReferencia = new Label();
			txtReferencia = new TextBox();
			btnReferencia = new Button();
			label3 = new Label();
			comboBoxSistema = new ComboBox();
			txtEstabelecimentoID = new TextBox();
			txtPessoaID = new TextBox();
			txtLoginID = new TextBox();
			label5 = new Label();
			listView1 = new ListView();
			btnAddToList = new Button();
			btnDelFromList = new Button();
			menuStrip1 = new MenuStrip();
			configuraçõesToolStripMenuItem = new ToolStripMenuItem();
			salvarNaPastaToolStripMenuItem = new ToolStripMenuItem();
			abrirPastaToolStripMenuItem = new ToolStripMenuItem();
			folderBrowserDialog1 = new FolderBrowserDialog();
			lbExcel2 = new Label();
			txtExcel2 = new TextBox();
			btnExcel2 = new Button();
			menuStrip1.SuspendLayout();
			SuspendLayout();
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Location = new Point(10, 67);
			label1.Name = "label1";
			label1.Size = new Size(108, 15);
			label1.TabIndex = 0;
			label1.Text = "EstabelecimentoID:";
			label1.Visible = false;
			// 
			// btnImportar
			// 
			btnImportar.Location = new Point(295, 231);
			btnImportar.Margin = new Padding(3, 2, 3, 2);
			btnImportar.Name = "btnImportar";
			btnImportar.Size = new Size(88, 22);
			btnImportar.TabIndex = 8;
			btnImportar.Text = "⚙ Executar";
			btnImportar.UseVisualStyleBackColor = true;
			btnImportar.Visible = false;
			btnImportar.Click += btnImportar_Click;
			// 
			// lbPessoaID
			// 
			lbPessoaID.AutoSize = true;
			lbPessoaID.Location = new Point(449, 67);
			lbPessoaID.Name = "lbPessoaID";
			lbPessoaID.Size = new Size(104, 15);
			lbPessoaID.TabIndex = 4;
			lbPessoaID.Text = "PessoaID Resp Fin:";
			lbPessoaID.Visible = false;
			// 
			// openFileDialog1
			// 
			openFileDialog1.FileName = "openFileDialog1";
			// 
			// btnExcel
			// 
			btnExcel.Location = new Point(638, 93);
			btnExcel.Margin = new Padding(3, 2, 3, 2);
			btnExcel.Name = "btnExcel";
			btnExcel.Size = new Size(31, 23);
			btnExcel.TabIndex = 5;
			btnExcel.Text = "📂";
			btnExcel.UseVisualStyleBackColor = true;
			btnExcel.Visible = false;
			btnExcel.Click += btnExcel_Click;
			// 
			// textBoxExcel1
			// 
			textBoxExcel1.Location = new Point(131, 94);
			textBoxExcel1.Margin = new Padding(3, 2, 3, 2);
			textBoxExcel1.Name = "textBoxExcel1";
			textBoxExcel1.Size = new Size(501, 23);
			textBoxExcel1.TabIndex = 90;
			textBoxExcel1.Visible = false;
			// 
			// labelExcel1
			// 
			labelExcel1.AutoSize = true;
			labelExcel1.Location = new Point(10, 97);
			labelExcel1.Name = "labelExcel1";
			labelExcel1.Size = new Size(43, 15);
			labelExcel1.TabIndex = 8;
			labelExcel1.Text = "Excel1:";
			labelExcel1.Visible = false;
			// 
			// comboBoxImportacao
			// 
			comboBoxImportacao.DropDownStyle = ComboBoxStyle.DropDownList;
			comboBoxImportacao.FormattingEnabled = true;
			comboBoxImportacao.Items.AddRange(new object[] { "JSON", "Fornecedores", "Pacientes", "Pagos", "Recebíveis", "Recebidos", "Tabela de Preços" });
			comboBoxImportacao.Location = new Point(131, 33);
			comboBoxImportacao.Margin = new Padding(3, 2, 3, 2);
			comboBoxImportacao.Name = "comboBoxImportacao";
			comboBoxImportacao.Size = new Size(214, 23);
			comboBoxImportacao.TabIndex = 0;
			comboBoxImportacao.SelectedIndexChanged += comboBoxImportacao_SelectedIndexChanged;
			// 
			// label4
			// 
			label4.AutoSize = true;
			label4.Location = new Point(12, 36);
			label4.Name = "label4";
			label4.Size = new Size(87, 15);
			label4.TabIndex = 10;
			label4.Text = "Importação de:";
			// 
			// lbReferencia
			// 
			lbReferencia.AutoSize = true;
			lbReferencia.Location = new Point(10, 151);
			lbReferencia.Name = "lbReferencia";
			lbReferencia.Size = new Size(65, 15);
			lbReferencia.TabIndex = 11;
			lbReferencia.Text = "Referência:";
			lbReferencia.Visible = false;
			// 
			// txtReferencia
			// 
			txtReferencia.Location = new Point(131, 148);
			txtReferencia.Margin = new Padding(3, 2, 3, 2);
			txtReferencia.Name = "txtReferencia";
			txtReferencia.Size = new Size(501, 23);
			txtReferencia.TabIndex = 91;
			txtReferencia.Visible = false;
			// 
			// btnReferencia
			// 
			btnReferencia.Location = new Point(638, 148);
			btnReferencia.Margin = new Padding(3, 2, 3, 2);
			btnReferencia.Name = "btnReferencia";
			btnReferencia.Size = new Size(31, 23);
			btnReferencia.TabIndex = 6;
			btnReferencia.Text = "📂";
			btnReferencia.UseVisualStyleBackColor = true;
			btnReferencia.Visible = false;
			btnReferencia.Click += btnReferencia_Click;
			// 
			// label3
			// 
			label3.AutoSize = true;
			label3.Location = new Point(353, 38);
			label3.Name = "label3";
			label3.Size = new Size(89, 15);
			label3.TabIndex = 15;
			label3.Text = "Antigo sistema:";
			label3.Visible = false;
			// 
			// comboBoxSistema
			// 
			comboBoxSistema.DropDownStyle = ComboBoxStyle.DropDownList;
			comboBoxSistema.FormattingEnabled = true;
			comboBoxSistema.Items.AddRange(new object[] { "DentalOffice", "OdontoCompany" });
			comboBoxSistema.Location = new Point(463, 34);
			comboBoxSistema.Margin = new Padding(3, 2, 3, 2);
			comboBoxSistema.Name = "comboBoxSistema";
			comboBoxSistema.Size = new Size(206, 23);
			comboBoxSistema.TabIndex = 1;
			comboBoxSistema.Visible = false;
			comboBoxSistema.SelectedIndexChanged += comboBoxSistema_SelectedIndexChanged;
			// 
			// txtEstabelecimentoID
			// 
			txtEstabelecimentoID.Location = new Point(131, 64);
			txtEstabelecimentoID.Margin = new Padding(3, 2, 3, 2);
			txtEstabelecimentoID.Name = "txtEstabelecimentoID";
			txtEstabelecimentoID.Size = new Size(113, 23);
			txtEstabelecimentoID.TabIndex = 2;
			txtEstabelecimentoID.Visible = false;
			txtEstabelecimentoID.TextChanged += txtEstabelecimentoID_TextChanged;
			txtEstabelecimentoID.KeyPress += txtEstabelecimentoID_KeyPress;
			// 
			// txtPessoaID
			// 
			txtPessoaID.Location = new Point(559, 64);
			txtPessoaID.Margin = new Padding(3, 2, 3, 2);
			txtPessoaID.Name = "txtPessoaID";
			txtPessoaID.Size = new Size(110, 23);
			txtPessoaID.TabIndex = 4;
			txtPessoaID.Visible = false;
			txtPessoaID.TextChanged += txtPessoaID_TextChanged;
			txtPessoaID.KeyPress += txtPessoaID_KeyPress;
			// 
			// txtLoginID
			// 
			txtLoginID.Location = new Point(323, 64);
			txtLoginID.Margin = new Padding(3, 2, 3, 2);
			txtLoginID.Name = "txtLoginID";
			txtLoginID.Size = new Size(110, 23);
			txtLoginID.TabIndex = 3;
			txtLoginID.Text = "1";
			txtLoginID.Visible = false;
			txtLoginID.KeyPress += txtLoginID_KeyPress;
			// 
			// label5
			// 
			label5.AutoSize = true;
			label5.Location = new Point(266, 67);
			label5.Name = "label5";
			label5.Size = new Size(51, 15);
			label5.TabIndex = 16;
			label5.Text = "LoginID:";
			label5.Visible = false;
			// 
			// listView1
			// 
			listView1.Location = new Point(12, 176);
			listView1.Name = "listView1";
			listView1.Size = new Size(620, 48);
			listView1.TabIndex = 18;
			listView1.UseCompatibleStateImageBehavior = false;
			listView1.View = View.List;
			listView1.Visible = false;
			// 
			// btnAddToList
			// 
			btnAddToList.Location = new Point(638, 175);
			btnAddToList.Margin = new Padding(3, 2, 3, 2);
			btnAddToList.Name = "btnAddToList";
			btnAddToList.Size = new Size(31, 23);
			btnAddToList.TabIndex = 7;
			btnAddToList.Text = "➕";
			btnAddToList.UseVisualStyleBackColor = true;
			btnAddToList.Visible = false;
			btnAddToList.Click += btnAddToList_Click;
			// 
			// btnDelFromList
			// 
			btnDelFromList.Location = new Point(638, 201);
			btnDelFromList.Margin = new Padding(3, 2, 3, 2);
			btnDelFromList.Name = "btnDelFromList";
			btnDelFromList.Size = new Size(31, 23);
			btnDelFromList.TabIndex = 20;
			btnDelFromList.Text = "🗑";
			btnDelFromList.UseVisualStyleBackColor = true;
			btnDelFromList.Visible = false;
			btnDelFromList.Click += btnDelFromList_Click;
			// 
			// menuStrip1
			// 
			menuStrip1.Items.AddRange(new ToolStripItem[] { configuraçõesToolStripMenuItem, abrirPastaToolStripMenuItem });
			menuStrip1.Location = new Point(0, 0);
			menuStrip1.Name = "menuStrip1";
			menuStrip1.Size = new Size(679, 24);
			menuStrip1.TabIndex = 95;
			menuStrip1.Text = "menuStrip1";
			// 
			// configuraçõesToolStripMenuItem
			// 
			configuraçõesToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { salvarNaPastaToolStripMenuItem });
			configuraçõesToolStripMenuItem.Name = "configuraçõesToolStripMenuItem";
			configuraçõesToolStripMenuItem.Size = new Size(96, 20);
			configuraçõesToolStripMenuItem.Text = "Configurações";
			// 
			// salvarNaPastaToolStripMenuItem
			// 
			salvarNaPastaToolStripMenuItem.Name = "salvarNaPastaToolStripMenuItem";
			salvarNaPastaToolStripMenuItem.Size = new Size(161, 22);
			salvarNaPastaToolStripMenuItem.Text = "Salvar na pasta...";
			salvarNaPastaToolStripMenuItem.Click += salvarNaPastaToolStripMenuItem_Click;
			// 
			// abrirPastaToolStripMenuItem
			// 
			abrirPastaToolStripMenuItem.Name = "abrirPastaToolStripMenuItem";
			abrirPastaToolStripMenuItem.Size = new Size(76, 20);
			abrirPastaToolStripMenuItem.Text = "Abrir Pasta";
			abrirPastaToolStripMenuItem.Click += abrirPastaToolStripMenuItem_Click;
			// 
			// lbExcel2
			// 
			lbExcel2.AutoSize = true;
			lbExcel2.Location = new Point(10, 124);
			lbExcel2.Name = "lbExcel2";
			lbExcel2.Size = new Size(43, 15);
			lbExcel2.TabIndex = 97;
			lbExcel2.Text = "Excel2:";
			lbExcel2.Visible = false;
			// 
			// txtExcel2
			// 
			txtExcel2.Location = new Point(131, 121);
			txtExcel2.Margin = new Padding(3, 2, 3, 2);
			txtExcel2.Name = "txtExcel2";
			txtExcel2.Size = new Size(501, 23);
			txtExcel2.TabIndex = 98;
			txtExcel2.Visible = false;
			// 
			// btnExcel2
			// 
			btnExcel2.Location = new Point(638, 120);
			btnExcel2.Margin = new Padding(3, 2, 3, 2);
			btnExcel2.Name = "btnExcel2";
			btnExcel2.Size = new Size(31, 23);
			btnExcel2.TabIndex = 96;
			btnExcel2.Text = "📂";
			btnExcel2.UseVisualStyleBackColor = true;
			btnExcel2.Visible = false;
			btnExcel2.Click += btnExcel2_Click_1;
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(679, 264);
			Controls.Add(lbExcel2);
			Controls.Add(txtExcel2);
			Controls.Add(btnExcel2);
			Controls.Add(btnDelFromList);
			Controls.Add(btnAddToList);
			Controls.Add(listView1);
			Controls.Add(txtLoginID);
			Controls.Add(label5);
			Controls.Add(txtPessoaID);
			Controls.Add(txtEstabelecimentoID);
			Controls.Add(label3);
			Controls.Add(comboBoxSistema);
			Controls.Add(btnReferencia);
			Controls.Add(txtReferencia);
			Controls.Add(lbReferencia);
			Controls.Add(label4);
			Controls.Add(comboBoxImportacao);
			Controls.Add(labelExcel1);
			Controls.Add(textBoxExcel1);
			Controls.Add(btnExcel);
			Controls.Add(lbPessoaID);
			Controls.Add(btnImportar);
			Controls.Add(label1);
			Controls.Add(menuStrip1);
			Icon = (Icon)resources.GetObject("$this.Icon");
			MainMenuStrip = menuStrip1;
			Margin = new Padding(3, 2, 3, 2);
			MaximizeBox = false;
			Name = "Form1";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "Migração";
			Load += Form1_Load;
			menuStrip1.ResumeLayout(false);
			menuStrip1.PerformLayout();
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label label1;
		private Button btnImportar;
		private Label lbPessoaID;
		private OpenFileDialog openFileDialog1;
		private Button btnExcel;
		private TextBox textBoxExcel1;
		private Label labelExcel1;
		private ComboBox comboBoxImportacao;
		private Label label4;
		private Label lbReferencia;
		private TextBox txtReferencia;
		private Button btnReferencia;
		private Label label3;
		private ComboBox comboBoxSistema;
		private TextBox txtEstabelecimentoID;
		private TextBox txtPessoaID;
		private TextBox txtLoginID;
		private Label label5;
		private ListView listView1;
		private Button btnAddToList;
		private Button btnDelFromList;
		private MenuStrip menuStrip1;
		private ToolStripMenuItem configuraçõesToolStripMenuItem;
		private ToolStripMenuItem salvarNaPastaToolStripMenuItem;
		private FolderBrowserDialog folderBrowserDialog1;
		private Label lbExcel2;
		private TextBox txtExcel2;
		private Button btnExcel2;
		private ToolStripMenuItem abrirPastaToolStripMenuItem;
	}
}
