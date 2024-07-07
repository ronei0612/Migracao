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
            openFileDialog1 = new OpenFileDialog();
            comboBoxImportacao = new ComboBox();
            label4 = new Label();
            label3 = new Label();
            comboBoxSistema = new ComboBox();
            txtEstabelecimentoID = new TextBox();
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
            toolStripMenuItem2 = new ToolStripMenuItem();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // lbEstabelecimento
            // 
            lbEstabelecimento.AutoSize = true;
            lbEstabelecimento.Location = new Point(12, 73);
            lbEstabelecimento.Name = "lbEstabelecimento";
            lbEstabelecimento.Size = new Size(108, 15);
            lbEstabelecimento.TabIndex = 0;
            lbEstabelecimento.Text = "EstabelecimentoID:";
            // 
            // btnImportar
            // 
            btnImportar.Anchor = AnchorStyles.Bottom;
            btnImportar.Location = new Point(365, 302);
            btnImportar.Margin = new Padding(3, 2, 3, 2);
            btnImportar.Name = "btnImportar";
            btnImportar.Size = new Size(88, 22);
            btnImportar.TabIndex = 8;
            btnImportar.Text = "✅ Executar";
            btnImportar.UseVisualStyleBackColor = true;
            btnImportar.Visible = false;
            btnImportar.Click += btnImportar_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // comboBoxImportacao
            // 
            comboBoxImportacao.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxImportacao.FormattingEnabled = true;
            comboBoxImportacao.Items.AddRange(new object[] { "JSON" });
            comboBoxImportacao.Location = new Point(133, 39);
            comboBoxImportacao.Margin = new Padding(3, 2, 3, 2);
            comboBoxImportacao.Name = "comboBoxImportacao";
            comboBoxImportacao.Size = new Size(214, 23);
            comboBoxImportacao.TabIndex = 0;
            comboBoxImportacao.SelectedIndexChanged += comboBoxImportacao_SelectedIndexChanged;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(12, 42);
            label4.Name = "label4";
            label4.Size = new Size(87, 15);
            label4.TabIndex = 10;
            label4.Text = "Importação de:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(355, 44);
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
            comboBoxSistema.Location = new Point(465, 40);
            comboBoxSistema.Margin = new Padding(3, 2, 3, 2);
            comboBoxSistema.Name = "comboBoxSistema";
            comboBoxSistema.Size = new Size(206, 23);
            comboBoxSistema.TabIndex = 1;
            comboBoxSistema.Visible = false;
            comboBoxSistema.SelectedIndexChanged += comboBoxSistema_SelectedIndexChanged;
            // 
            // txtEstabelecimentoID
            // 
            txtEstabelecimentoID.Location = new Point(133, 70);
            txtEstabelecimentoID.Margin = new Padding(3, 2, 3, 2);
            txtEstabelecimentoID.Name = "txtEstabelecimentoID";
            txtEstabelecimentoID.Size = new Size(113, 23);
            txtEstabelecimentoID.TabIndex = 2;
            txtEstabelecimentoID.KeyPress += txtEstabelecimentoID_KeyPress;
            txtEstabelecimentoID.Leave += txtEstabelecimentoID_Leave;
            // 
            // listView1
            // 
            listView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listView1.Columns.AddRange(new ColumnHeader[] { columnHeader1 });
            listView1.Location = new Point(12, 98);
            listView1.Name = "listView1";
            listView1.Size = new Size(759, 199);
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
            btnAddToList.Location = new Point(777, 98);
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
            btnDelFromList.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnDelFromList.Location = new Point(777, 124);
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
            menuStrip1.ImageScalingSize = new Size(20, 20);
            menuStrip1.Items.AddRange(new ToolStripItem[] { configuraçõesToolStripMenuItem, abrirPastaToolStripMenuItem, importarDataBaseToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(818, 24);
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
            // importarDataBaseToolStripMenuItem
            // 
            importarDataBaseToolStripMenuItem.Name = "importarDataBaseToolStripMenuItem";
            importarDataBaseToolStripMenuItem.Size = new Size(116, 20);
            importarDataBaseToolStripMenuItem.Text = "Importar DataBase";
            importarDataBaseToolStripMenuItem.Click += OpenFormImportarDataBase;
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(224, 26);
            toolStripMenuItem2.Text = "Salvar na pasta...";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(818, 335);
            Controls.Add(btnDelFromList);
            Controls.Add(btnAddToList);
            Controls.Add(listView1);
            Controls.Add(txtEstabelecimentoID);
            Controls.Add(label3);
            Controls.Add(comboBoxSistema);
            Controls.Add(label4);
            Controls.Add(comboBoxImportacao);
            Controls.Add(btnImportar);
            Controls.Add(lbEstabelecimento);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Margin = new Padding(3, 2, 3, 2);
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
		private OpenFileDialog openFileDialog1;
		private ComboBox comboBoxImportacao;
		private Label label4;
		private Label label3;
		private ComboBox comboBoxSistema;
		private TextBox txtEstabelecimentoID;
		private ListView listView1;
		private Button btnAddToList;
		private Button btnDelFromList;
		private MenuStrip menuStrip1;
		private ToolStripMenuItem configuraçõesToolStripMenuItem;
		private ToolStripMenuItem salvarNaPastaToolStripMenuItem;
		private FolderBrowserDialog folderBrowserDialog1;
		private ToolStripMenuItem abrirPastaToolStripMenuItem;
		private ColumnHeader columnHeader1;
        private ToolStripMenuItem btnImportarDataBase;
        private ToolStripMenuItem toolStripMenuItem2;
        private ToolStripMenuItem importarDataBaseToolStripMenuItem;
    }
}
