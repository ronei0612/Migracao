namespace Migracao
{
    partial class ImportacaoDataBase
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
            inputDB = new TextBox();
            labelPathDB = new Label();
            labelPathDBContratos = new Label();
            inputDBContratos = new TextBox();
            BtnPathDB = new Button();
            BtnPathDBContratos = new Button();
            label3 = new Label();
            comboBoxSistema = new ComboBox();
            panelDBContratos = new Panel();
            panelDB = new Panel();
            panelAntigoSistema = new Panel();
            comboTabelas = new ComboBox();
            label1 = new Label();
            btnImportar = new Button();
            panelDataBase = new Panel();
            inputDataBaseName = new TextBox();
            labelDataBase = new Label();
            panelDBContratos.SuspendLayout();
            panelDB.SuspendLayout();
            panelAntigoSistema.SuspendLayout();
            panelDataBase.SuspendLayout();
            SuspendLayout();
            // 
            // inputDB
            // 
            inputDB.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            inputDB.Enabled = false;
            inputDB.Location = new Point(186, 8);
            inputDB.Margin = new Padding(3, 2, 3, 2);
            inputDB.Name = "inputDB";
            inputDB.Size = new Size(532, 23);
            inputDB.TabIndex = 102;
            // 
            // labelPathDB
            // 
            labelPathDB.AutoSize = true;
            labelPathDB.Location = new Point(3, 10);
            labelPathDB.Name = "labelPathDB";
            labelPathDB.Size = new Size(146, 15);
            labelPathDB.TabIndex = 103;
            labelPathDB.Text = "Caminho DataBase Clinica";
            // 
            // labelPathDBContratos
            // 
            labelPathDBContratos.AutoSize = true;
            labelPathDBContratos.Location = new Point(3, 11);
            labelPathDBContratos.Name = "labelPathDBContratos";
            labelPathDBContratos.Size = new Size(162, 15);
            labelPathDBContratos.TabIndex = 105;
            labelPathDBContratos.Text = "Caminho DataBase Contratos";
            // 
            // inputDBContratos
            // 
            inputDBContratos.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            inputDBContratos.Enabled = false;
            inputDBContratos.Location = new Point(186, 9);
            inputDBContratos.Margin = new Padding(3, 2, 3, 2);
            inputDBContratos.Name = "inputDBContratos";
            inputDBContratos.Size = new Size(532, 23);
            inputDBContratos.TabIndex = 104;
            // 
            // BtnPathDB
            // 
            BtnPathDB.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            BtnPathDB.Location = new Point(735, 8);
            BtnPathDB.Margin = new Padding(3, 2, 3, 2);
            BtnPathDB.Name = "BtnPathDB";
            BtnPathDB.Size = new Size(31, 23);
            BtnPathDB.TabIndex = 106;
            BtnPathDB.Text = "📂";
            BtnPathDB.UseVisualStyleBackColor = true;
            BtnPathDB.Click += BtnPathDB_Click;
            BtnPathDB.KeyDown += BtnPathDB_KeyDown;
            // 
            // BtnPathDBContratos
            // 
            BtnPathDBContratos.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            BtnPathDBContratos.Location = new Point(735, 8);
            BtnPathDBContratos.Margin = new Padding(3, 2, 3, 2);
            BtnPathDBContratos.Name = "BtnPathDBContratos";
            BtnPathDBContratos.Size = new Size(31, 23);
            BtnPathDBContratos.TabIndex = 107;
            BtnPathDBContratos.Text = "📂";
            BtnPathDBContratos.UseVisualStyleBackColor = true;
            BtnPathDBContratos.Click += BtnPathDBContratos_Click;
            BtnPathDBContratos.KeyDown += BtnPathDB_KeyDown;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(8, 6);
            label3.Name = "label3";
            label3.Size = new Size(89, 15);
            label3.TabIndex = 109;
            label3.Text = "Antigo sistema:";
            // 
            // comboBoxSistema
            // 
            comboBoxSistema.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxSistema.FormattingEnabled = true;
            comboBoxSistema.Items.AddRange(new object[] { "DentalOffice", "OdontoCompany" });
            comboBoxSistema.Location = new Point(122, 4);
            comboBoxSistema.Margin = new Padding(3, 2, 3, 2);
            comboBoxSistema.Name = "comboBoxSistema";
            comboBoxSistema.Size = new Size(206, 23);
            comboBoxSistema.TabIndex = 108;
            comboBoxSistema.SelectedIndexChanged += AntigoSistemaChanged;
            comboBoxSistema.KeyDown += BtnPathDB_KeyDown;
            // 
            // panelDBContratos
            // 
            panelDBContratos.Controls.Add(inputDBContratos);
            panelDBContratos.Controls.Add(labelPathDBContratos);
            panelDBContratos.Controls.Add(BtnPathDBContratos);
            panelDBContratos.Location = new Point(27, 99);
            panelDBContratos.Margin = new Padding(3, 2, 3, 2);
            panelDBContratos.Name = "panelDBContratos";
            panelDBContratos.Size = new Size(768, 43);
            panelDBContratos.TabIndex = 110;
            panelDBContratos.Visible = false;
            // 
            // panelDB
            // 
            panelDB.Controls.Add(inputDB);
            panelDB.Controls.Add(BtnPathDB);
            panelDB.Controls.Add(labelPathDB);
            panelDB.Location = new Point(27, 57);
            panelDB.Margin = new Padding(3, 2, 3, 2);
            panelDB.Name = "panelDB";
            panelDB.Size = new Size(768, 38);
            panelDB.TabIndex = 111;
            panelDB.Visible = false;
            // 
            // panelAntigoSistema
            // 
            panelAntigoSistema.Controls.Add(label3);
            panelAntigoSistema.Controls.Add(comboBoxSistema);
            panelAntigoSistema.Location = new Point(249, 9);
            panelAntigoSistema.Margin = new Padding(3, 2, 3, 2);
            panelAntigoSistema.Name = "panelAntigoSistema";
            panelAntigoSistema.Size = new Size(340, 32);
            panelAntigoSistema.TabIndex = 112;
            // 
            // comboTabelas
            // 
            comboTabelas.DropDownStyle = ComboBoxStyle.DropDownList;
            comboTabelas.FormattingEnabled = true;
            comboTabelas.Items.AddRange(new object[] { "TUDO", "PacientesDentistas", "AgendamentosDesenvClínico", "ProcedimentosTabela Preços", "Recebíveis Pagos e Exigíveis", "ProcedimentosManutenções" });
            comboTabelas.Location = new Point(361, 196);
            comboTabelas.Margin = new Padding(3, 2, 3, 2);
            comboTabelas.Name = "comboTabelas";
            comboTabelas.Size = new Size(206, 23);
            comboTabelas.TabIndex = 113;
            comboTabelas.KeyDown += BtnPathDB_KeyDown;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(293, 199);
            label1.Name = "label1";
            label1.Size = new Size(56, 15);
            label1.TabIndex = 114;
            label1.Text = "Importar:";
            // 
            // btnImportar
            // 
            btnImportar.Anchor = AnchorStyles.Bottom;
            btnImportar.Location = new Point(372, 358);
            btnImportar.Margin = new Padding(3, 2, 3, 2);
            btnImportar.Name = "btnImportar";
            btnImportar.Size = new Size(113, 31);
            btnImportar.TabIndex = 115;
            btnImportar.Text = "✅ Executar";
            btnImportar.UseVisualStyleBackColor = true;
            btnImportar.Click += btnImportar_Click;
            btnImportar.KeyDown += BtnPathDB_KeyDown;
            // 
            // panelDataBase
            // 
            panelDataBase.Controls.Add(inputDataBaseName);
            panelDataBase.Controls.Add(labelDataBase);
            panelDataBase.Location = new Point(257, 155);
            panelDataBase.Margin = new Padding(3, 2, 3, 2);
            panelDataBase.Name = "panelDataBase";
            panelDataBase.Size = new Size(332, 37);
            panelDataBase.TabIndex = 116;
            // 
            // inputDataBaseName
            // 
            inputDataBaseName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            inputDataBaseName.Location = new Point(110, 8);
            inputDataBaseName.Margin = new Padding(3, 2, 3, 2);
            inputDataBaseName.Name = "inputDataBaseName";
            inputDataBaseName.Size = new Size(210, 23);
            inputDataBaseName.TabIndex = 105;
            // 
            // labelDataBase
            // 
            labelDataBase.AutoSize = true;
            labelDataBase.Location = new Point(3, 10);
            labelDataBase.Name = "labelDataBase";
            labelDataBase.Size = new Size(91, 15);
            labelDataBase.TabIndex = 104;
            labelDataBase.Text = "DataBase Nome";
            // 
            // ImportacaoDataBase
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(836, 411);
            Controls.Add(panelDataBase);
            Controls.Add(btnImportar);
            Controls.Add(label1);
            Controls.Add(comboTabelas);
            Controls.Add(panelDBContratos);
            Controls.Add(panelDB);
            Controls.Add(panelAntigoSistema);
            Margin = new Padding(3, 2, 3, 2);
            Name = "ImportacaoDataBase";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Importação DataBase";
            Load += ImportacaoDataBase_Load;
            KeyDown += ImportacaoDataBase_KeyDown;
            panelDBContratos.ResumeLayout(false);
            panelDBContratos.PerformLayout();
            panelDB.ResumeLayout(false);
            panelDB.PerformLayout();
            panelAntigoSistema.ResumeLayout(false);
            panelAntigoSistema.PerformLayout();
            panelDataBase.ResumeLayout(false);
            panelDataBase.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox inputDB;
        private Label labelPathDB;
        private Label labelPathDBContratos;
        private TextBox inputDBContratos;
        private Button BtnPathDB;
        private Button BtnPathDBContratos;
        private Label label3;
        private ComboBox comboBoxSistema;
        private Panel panelDBContratos;
        private Panel panelDB;
        private Panel panelAntigoSistema;
        private ComboBox comboTabelas;
        private Label label1;
        private Button btnImportar;
        private Panel panelDataBase;
        private TextBox inputDataBaseName;
        private Label labelDataBase;
    }
}