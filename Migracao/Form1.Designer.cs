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
            lbEstabelecimento = new Label();
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
            columnHeader1 = new ColumnHeader();
            btnAddToList = new Button();
            btnDelFromList = new Button();
            menuStrip1 = new MenuStrip();
            configuraçõesToolStripMenuItem = new ToolStripMenuItem();
            salvarNaPastaToolStripMenuItem = new ToolStripMenuItem();
            abrirPastaToolStripMenuItem = new ToolStripMenuItem();
            importarDataBaseToolStripMenuItem = new ToolStripMenuItem();
            folderBrowserDialog1 = new FolderBrowserDialog();
            lbExcel2 = new Label();
            txtExcel2 = new TextBox();
            btnExcel2 = new Button();
            label2 = new Label();
            label6 = new Label();
            txtPessoas = new TextBox();
            btnPessoas = new Button();
            txtRecebiveis = new TextBox();
            btnRecebiveis = new Button();
            toolStripMenuItem2 = new ToolStripMenuItem();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // lbEstabelecimento
            // 
            lbEstabelecimento.AutoSize = true;
            lbEstabelecimento.Location = new Point(11, 167);
            lbEstabelecimento.Name = "lbEstabelecimento";
            lbEstabelecimento.Size = new Size(137, 20);
            lbEstabelecimento.TabIndex = 0;
            lbEstabelecimento.Text = "EstabelecimentoID:";
            // 
            // btnImportar
            // 
            btnImportar.Anchor = AnchorStyles.Bottom;
            btnImportar.Location = new Point(416, 576);
            btnImportar.Name = "btnImportar";
            btnImportar.Size = new Size(101, 29);
            btnImportar.TabIndex = 8;
            btnImportar.Text = "⚙ Executar";
            btnImportar.UseVisualStyleBackColor = true;
            btnImportar.Visible = false;
            btnImportar.Click += btnImportar_Click;
            // 
            // lbPessoaID
            // 
            lbPessoaID.AutoSize = true;
            lbPessoaID.Location = new Point(517, 167);
            lbPessoaID.Name = "lbPessoaID";
            lbPessoaID.Size = new Size(126, 20);
            lbPessoaID.TabIndex = 4;
            lbPessoaID.Text = "PessoaID RespFin:";
            lbPessoaID.Visible = false;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnExcel
            // 
            btnExcel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnExcel.Location = new Point(888, 201);
            btnExcel.Name = "btnExcel";
            btnExcel.Size = new Size(35, 31);
            btnExcel.TabIndex = 5;
            btnExcel.Text = "📂";
            btnExcel.UseVisualStyleBackColor = true;
            btnExcel.Visible = false;
            btnExcel.Click += btnExcel_Click;
            // 
            // textBoxExcel1
            // 
            textBoxExcel1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            textBoxExcel1.Location = new Point(150, 203);
            textBoxExcel1.Name = "textBoxExcel1";
            textBoxExcel1.Size = new Size(731, 27);
            textBoxExcel1.TabIndex = 90;
            textBoxExcel1.Visible = false;
            // 
            // labelExcel1
            // 
            labelExcel1.AutoSize = true;
            labelExcel1.Location = new Point(11, 207);
            labelExcel1.Name = "labelExcel1";
            labelExcel1.Size = new Size(54, 20);
            labelExcel1.TabIndex = 8;
            labelExcel1.Text = "Excel1:";
            labelExcel1.Visible = false;
            // 
            // comboBoxImportacao
            // 
            comboBoxImportacao.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxImportacao.FormattingEnabled = true;
            comboBoxImportacao.Items.AddRange(new object[] { "JSON", "Todos", "Agendamentos", "Fornecedores", "Pessoas", "Pagos", "Recebíveis", "Preços", "Desenv Clinico" });
            comboBoxImportacao.Location = new Point(150, 121);
            comboBoxImportacao.Name = "comboBoxImportacao";
            comboBoxImportacao.Size = new Size(244, 28);
            comboBoxImportacao.TabIndex = 0;
            comboBoxImportacao.SelectedIndexChanged += comboBoxImportacao_SelectedIndexChanged;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(11, 125);
            label4.Name = "label4";
            label4.Size = new Size(110, 20);
            label4.TabIndex = 10;
            label4.Text = "Importação de:";
            // 
            // lbReferencia
            // 
            lbReferencia.AutoSize = true;
            lbReferencia.Location = new Point(11, 279);
            lbReferencia.Name = "lbReferencia";
            lbReferencia.Size = new Size(82, 20);
            lbReferencia.TabIndex = 11;
            lbReferencia.Text = "Referência:";
            lbReferencia.Visible = false;
            // 
            // txtReferencia
            // 
            txtReferencia.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtReferencia.Location = new Point(150, 275);
            txtReferencia.Name = "txtReferencia";
            txtReferencia.Size = new Size(731, 27);
            txtReferencia.TabIndex = 91;
            txtReferencia.Visible = false;
            // 
            // btnReferencia
            // 
            btnReferencia.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnReferencia.Location = new Point(888, 275);
            btnReferencia.Name = "btnReferencia";
            btnReferencia.Size = new Size(35, 31);
            btnReferencia.TabIndex = 6;
            btnReferencia.Text = "📂";
            btnReferencia.UseVisualStyleBackColor = true;
            btnReferencia.Visible = false;
            btnReferencia.Click += btnReferencia_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(403, 128);
            label3.Name = "label3";
            label3.Size = new Size(111, 20);
            label3.TabIndex = 15;
            label3.Text = "Antigo sistema:";
            label3.Visible = false;
            // 
            // comboBoxSistema
            // 
            comboBoxSistema.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxSistema.FormattingEnabled = true;
            comboBoxSistema.Items.AddRange(new object[] { "DentalOffice", "OdontoCompany" });
            comboBoxSistema.Location = new Point(529, 123);
            comboBoxSistema.Name = "comboBoxSistema";
            comboBoxSistema.Size = new Size(235, 28);
            comboBoxSistema.TabIndex = 1;
            comboBoxSistema.Visible = false;
            comboBoxSistema.SelectedIndexChanged += comboBoxSistema_SelectedIndexChanged;
            // 
            // txtEstabelecimentoID
            // 
            txtEstabelecimentoID.Location = new Point(150, 163);
            txtEstabelecimentoID.Name = "txtEstabelecimentoID";
            txtEstabelecimentoID.Size = new Size(129, 27);
            txtEstabelecimentoID.TabIndex = 2;
            txtEstabelecimentoID.KeyPress += txtEstabelecimentoID_KeyPress;
            // 
            // txtPessoaID
            // 
            txtPessoaID.Location = new Point(639, 163);
            txtPessoaID.Name = "txtPessoaID";
            txtPessoaID.Size = new Size(125, 27);
            txtPessoaID.TabIndex = 4;
            txtPessoaID.Visible = false;
            txtPessoaID.KeyPress += txtPessoaID_KeyPress;
            // 
            // txtLoginID
            // 
            txtLoginID.Location = new Point(369, 163);
            txtLoginID.Name = "txtLoginID";
            txtLoginID.Size = new Size(125, 27);
            txtLoginID.TabIndex = 3;
            txtLoginID.Text = "1";
            txtLoginID.Visible = false;
            txtLoginID.KeyPress += txtLoginID_KeyPress;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(304, 167);
            label5.Name = "label5";
            label5.Size = new Size(64, 20);
            label5.TabIndex = 16;
            label5.Text = "LoginID:";
            label5.Visible = false;
            // 
            // listView1
            // 
            listView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listView1.Columns.AddRange(new ColumnHeader[] { columnHeader1 });
            listView1.Location = new Point(14, 311);
            listView1.Margin = new Padding(3, 4, 3, 4);
            listView1.Name = "listView1";
            listView1.Size = new Size(867, 257);
            listView1.TabIndex = 18;
            listView1.UseCompatibleStateImageBehavior = false;
            listView1.View = View.Details;
            listView1.Visible = false;
            // 
            // columnHeader1
            // 
            columnHeader1.Text = "";
            columnHeader1.Width = 2000;
            // 
            // btnAddToList
            // 
            btnAddToList.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnAddToList.Location = new Point(888, 311);
            btnAddToList.Name = "btnAddToList";
            btnAddToList.Size = new Size(35, 31);
            btnAddToList.TabIndex = 7;
            btnAddToList.Text = "➕";
            btnAddToList.UseVisualStyleBackColor = true;
            btnAddToList.Visible = false;
            btnAddToList.Click += btnAddToList_Click;
            // 
            // btnDelFromList
            // 
            btnDelFromList.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnDelFromList.Location = new Point(888, 345);
            btnDelFromList.Name = "btnDelFromList";
            btnDelFromList.Size = new Size(35, 31);
            btnDelFromList.TabIndex = 20;
            btnDelFromList.Text = "🗑";
            btnDelFromList.UseVisualStyleBackColor = true;
            btnDelFromList.Visible = false;
            btnDelFromList.Click += btnDelFromList_Click;
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(20, 20);
            menuStrip1.Items.AddRange(new ToolStripItem[] { configuraçõesToolStripMenuItem, abrirPastaToolStripMenuItem, importarDataBaseToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Padding = new Padding(7, 3, 0, 3);
            menuStrip1.Size = new Size(935, 30);
            menuStrip1.TabIndex = 95;
            menuStrip1.Text = "menuStrip1";
            // 
            // configuraçõesToolStripMenuItem
            // 
            configuraçõesToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { salvarNaPastaToolStripMenuItem });
            configuraçõesToolStripMenuItem.Name = "configuraçõesToolStripMenuItem";
            configuraçõesToolStripMenuItem.Size = new Size(118, 24);
            configuraçõesToolStripMenuItem.Text = "Configurações";
            // 
            // salvarNaPastaToolStripMenuItem
            // 
            salvarNaPastaToolStripMenuItem.Name = "salvarNaPastaToolStripMenuItem";
            salvarNaPastaToolStripMenuItem.Size = new Size(201, 26);
            salvarNaPastaToolStripMenuItem.Text = "Salvar na pasta...";
            salvarNaPastaToolStripMenuItem.Click += salvarNaPastaToolStripMenuItem_Click;
            // 
            // abrirPastaToolStripMenuItem
            // 
            abrirPastaToolStripMenuItem.Name = "abrirPastaToolStripMenuItem";
            abrirPastaToolStripMenuItem.Size = new Size(94, 24);
            abrirPastaToolStripMenuItem.Text = "Abrir Pasta";
            abrirPastaToolStripMenuItem.Click += abrirPastaToolStripMenuItem_Click;
            // 
            // importarDataBaseToolStripMenuItem
            // 
            importarDataBaseToolStripMenuItem.Name = "importarDataBaseToolStripMenuItem";
            importarDataBaseToolStripMenuItem.Size = new Size(148, 24);
            importarDataBaseToolStripMenuItem.Text = "Importar DataBase";
            importarDataBaseToolStripMenuItem.Click += OpenFormImportarDataBase;
            // 
            // lbExcel2
            // 
            lbExcel2.AutoSize = true;
            lbExcel2.Location = new Point(11, 243);
            lbExcel2.Name = "lbExcel2";
            lbExcel2.Size = new Size(54, 20);
            lbExcel2.TabIndex = 97;
            lbExcel2.Text = "Excel2:";
            lbExcel2.Visible = false;
            // 
            // txtExcel2
            // 
            txtExcel2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtExcel2.Location = new Point(150, 239);
            txtExcel2.Name = "txtExcel2";
            txtExcel2.Size = new Size(731, 27);
            txtExcel2.TabIndex = 98;
            txtExcel2.Visible = false;
            // 
            // btnExcel2
            // 
            btnExcel2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnExcel2.Location = new Point(888, 237);
            btnExcel2.Name = "btnExcel2";
            btnExcel2.Size = new Size(35, 31);
            btnExcel2.TabIndex = 96;
            btnExcel2.Text = "📂";
            btnExcel2.UseVisualStyleBackColor = true;
            btnExcel2.Visible = false;
            btnExcel2.Click += btnExcel2_Click_1;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(14, 39);
            label2.Name = "label2";
            label2.Size = new Size(62, 20);
            label2.TabIndex = 99;
            label2.Text = "Pessoas:";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(11, 79);
            label6.Name = "label6";
            label6.Size = new Size(82, 20);
            label6.TabIndex = 100;
            label6.Text = "Recebíveis:";
            // 
            // txtPessoas
            // 
            txtPessoas.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtPessoas.Enabled = false;
            txtPessoas.Location = new Point(149, 35);
            txtPessoas.Name = "txtPessoas";
            txtPessoas.Size = new Size(731, 27);
            txtPessoas.TabIndex = 101;
            // 
            // btnPessoas
            // 
            btnPessoas.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnPessoas.Location = new Point(886, 35);
            btnPessoas.Name = "btnPessoas";
            btnPessoas.Size = new Size(35, 31);
            btnPessoas.TabIndex = 102;
            btnPessoas.Text = "📂";
            btnPessoas.UseVisualStyleBackColor = true;
            btnPessoas.Click += btnPessoas_Click;
            // 
            // txtRecebiveis
            // 
            txtRecebiveis.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtRecebiveis.Enabled = false;
            txtRecebiveis.Location = new Point(149, 75);
            txtRecebiveis.Name = "txtRecebiveis";
            txtRecebiveis.Size = new Size(731, 27);
            txtRecebiveis.TabIndex = 103;
            // 
            // btnRecebiveis
            // 
            btnRecebiveis.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnRecebiveis.Location = new Point(886, 75);
            btnRecebiveis.Name = "btnRecebiveis";
            btnRecebiveis.Size = new Size(35, 31);
            btnRecebiveis.TabIndex = 104;
            btnRecebiveis.Text = "📂";
            btnRecebiveis.UseVisualStyleBackColor = true;
            btnRecebiveis.Click += btnRecebiveis_Click;
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(224, 26);
            toolStripMenuItem2.Text = "Salvar na pasta...";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(935, 618);
            Controls.Add(btnRecebiveis);
            Controls.Add(txtRecebiveis);
            Controls.Add(btnPessoas);
            Controls.Add(txtPessoas);
            Controls.Add(label6);
            Controls.Add(label2);
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
            Controls.Add(lbEstabelecimento);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
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

        private Label lbEstabelecimento;
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
		private Label label2;
		private Label label6;
		private TextBox txtPessoas;
		private Button btnPessoas;
		private TextBox txtRecebiveis;
		private Button btnRecebiveis;
		private ColumnHeader columnHeader1;
        private ToolStripMenuItem btnImportarDataBase;
        private ToolStripMenuItem toolStripMenuItem2;
        private ToolStripMenuItem importarDataBaseToolStripMenuItem;
    }
}
