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
            inputDB.Location = new Point(212, 10);
            inputDB.Name = "inputDB";
            inputDB.Size = new Size(608, 27);
            inputDB.TabIndex = 102;
            // 
            // labelPathDB
            // 
            labelPathDB.AutoSize = true;
            labelPathDB.Location = new Point(3, 13);
            labelPathDB.Name = "labelPathDB";
            labelPathDB.Size = new Size(183, 20);
            labelPathDB.TabIndex = 103;
            labelPathDB.Text = "Caminho DataBase Clinica";
            labelPathDB.Click += label2_Click;
            // 
            // labelPathDBContratos
            // 
            labelPathDBContratos.AutoSize = true;
            labelPathDBContratos.Location = new Point(3, 15);
            labelPathDBContratos.Name = "labelPathDBContratos";
            labelPathDBContratos.Size = new Size(203, 20);
            labelPathDBContratos.TabIndex = 105;
            labelPathDBContratos.Text = "Caminho DataBase Contratos";
            // 
            // inputDBContratos
            // 
            inputDBContratos.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            inputDBContratos.Enabled = false;
            inputDBContratos.Location = new Point(212, 12);
            inputDBContratos.Name = "inputDBContratos";
            inputDBContratos.Size = new Size(608, 27);
            inputDBContratos.TabIndex = 104;
            // 
            // BtnPathDB
            // 
            BtnPathDB.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            BtnPathDB.Location = new Point(840, 10);
            BtnPathDB.Name = "BtnPathDB";
            BtnPathDB.Size = new Size(35, 31);
            BtnPathDB.TabIndex = 106;
            BtnPathDB.Text = "📂";
            BtnPathDB.UseVisualStyleBackColor = true;
            BtnPathDB.Click += BtnPathDB_Click;
            // 
            // BtnPathDBContratos
            // 
            BtnPathDBContratos.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            BtnPathDBContratos.Location = new Point(840, 10);
            BtnPathDBContratos.Name = "BtnPathDBContratos";
            BtnPathDBContratos.Size = new Size(35, 31);
            BtnPathDBContratos.TabIndex = 107;
            BtnPathDBContratos.Text = "📂";
            BtnPathDBContratos.UseVisualStyleBackColor = true;
            BtnPathDBContratos.Click += BtnPathDBContratos_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(9, 8);
            label3.Name = "label3";
            label3.Size = new Size(111, 20);
            label3.TabIndex = 109;
            label3.Text = "Antigo sistema:";
            // 
            // comboBoxSistema
            // 
            comboBoxSistema.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxSistema.FormattingEnabled = true;
            comboBoxSistema.Items.AddRange(new object[] { "DentalOffice", "OdontoCompany" });
            comboBoxSistema.Location = new Point(140, 5);
            comboBoxSistema.Name = "comboBoxSistema";
            comboBoxSistema.Size = new Size(235, 28);
            comboBoxSistema.TabIndex = 108;
            comboBoxSistema.SelectedIndexChanged += AntigoSistemaChanged;
            // 
            // panelDBContratos
            // 
            panelDBContratos.Controls.Add(inputDBContratos);
            panelDBContratos.Controls.Add(labelPathDBContratos);
            panelDBContratos.Controls.Add(BtnPathDBContratos);
            panelDBContratos.Location = new Point(31, 132);
            panelDBContratos.Name = "panelDBContratos";
            panelDBContratos.Size = new Size(878, 57);
            panelDBContratos.TabIndex = 110;
            panelDBContratos.Visible = false;
            // 
            // panelDB
            // 
            panelDB.Controls.Add(inputDB);
            panelDB.Controls.Add(BtnPathDB);
            panelDB.Controls.Add(labelPathDB);
            panelDB.Location = new Point(31, 76);
            panelDB.Name = "panelDB";
            panelDB.Size = new Size(878, 50);
            panelDB.TabIndex = 111;
            panelDB.Visible = false;
            // 
            // panelAntigoSistema
            // 
            panelAntigoSistema.Controls.Add(label3);
            panelAntigoSistema.Controls.Add(comboBoxSistema);
            panelAntigoSistema.Location = new Point(285, 12);
            panelAntigoSistema.Name = "panelAntigoSistema";
            panelAntigoSistema.Size = new Size(388, 43);
            panelAntigoSistema.TabIndex = 112;
            // 
            // comboTabelas
            // 
            comboTabelas.DropDownStyle = ComboBoxStyle.DropDownList;
            comboTabelas.FormattingEnabled = true;
            comboTabelas.Items.AddRange(new object[] { "", "Prontuários", "Agendamentos", "Desenvolvimento Clínico", "Procedimentos", "Manutenções", "Financeiro (Receber)" });
            comboTabelas.Location = new Point(413, 262);
            comboTabelas.Name = "comboTabelas";
            comboTabelas.Size = new Size(235, 28);
            comboTabelas.TabIndex = 113;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(335, 265);
            label1.Name = "label1";
            label1.Size = new Size(70, 20);
            label1.TabIndex = 114;
            label1.Text = "Importar:";
            // 
            // btnImportar
            // 
            btnImportar.Anchor = AnchorStyles.Bottom;
            btnImportar.Location = new Point(425, 478);
            btnImportar.Name = "btnImportar";
            btnImportar.Size = new Size(129, 41);
            btnImportar.TabIndex = 115;
            btnImportar.Text = "⚙ Executar";
            btnImportar.UseVisualStyleBackColor = true;
            btnImportar.Click += ExecutarImportacao;
            // 
            // panelDataBase
            // 
            panelDataBase.Controls.Add(inputDataBaseName);
            panelDataBase.Controls.Add(labelDataBase);
            panelDataBase.Location = new Point(294, 207);
            panelDataBase.Name = "panelDataBase";
            panelDataBase.Size = new Size(379, 49);
            panelDataBase.TabIndex = 116;
            // 
            // inputDataBaseName
            // 
            inputDataBaseName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            inputDataBaseName.Location = new Point(126, 11);
            inputDataBaseName.Name = "inputDataBaseName";
            inputDataBaseName.Size = new Size(240, 27);
            inputDataBaseName.TabIndex = 105;
            // 
            // labelDataBase
            // 
            labelDataBase.AutoSize = true;
            labelDataBase.Location = new Point(3, 14);
            labelDataBase.Name = "labelDataBase";
            labelDataBase.Size = new Size(117, 20);
            labelDataBase.TabIndex = 104;
            labelDataBase.Text = "DataBase Nome";
            // 
            // ImportacaoDataBase
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(955, 548);
            Controls.Add(panelDataBase);
            Controls.Add(btnImportar);
            Controls.Add(label1);
            Controls.Add(comboTabelas);
            Controls.Add(panelDBContratos);
            Controls.Add(panelDB);
            Controls.Add(panelAntigoSistema);
            Name = "ImportacaoDataBase";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Importação DataBase";
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