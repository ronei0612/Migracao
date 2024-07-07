using Migracao.Imports;
using Migracao.Sistems;
using Migracao.Utils;

namespace Migracao
{
    public partial class Form1 : Form
    {
        string nomeArquivoExcel = "";
        string janelaArquivoExcel = "Selecione um arquivo";

        public Form1()
        {
            InitializeComponent();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                Filter = "Arquivo Excel |*.xlsx",
                Title = janelaArquivoExcel,
                FileName = nomeArquivoExcel,
                InitialDirectory = Tools.ultimaPasta
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
                Tools.SalvarConfig();
            }
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            Tools.ultimoEstabelecimentoID = txtEstabelecimentoID.Text;
            Tools.ultimoAntigoSistema = comboBoxSistema.SelectedIndex.ToString();
            Tools.SalvarConfig();

            if (ValidarCampos())
            {
                try
                {
                    if (listView1.Visible == true)
                    {
                        if (comboBoxImportacao.Text.Equals("json", StringComparison.CurrentCultureIgnoreCase))
                        {
                            ConverterHelper converterHelper = new ConverterHelper();
                            var nomeArquivo = "";

                            foreach (ListViewItem item in listView1.Items)
                            {
                                nomeArquivo = Tools.TratarCaracteres(Path.GetFileNameWithoutExtension(item.Text));
                                var pastaArquivo = Path.GetDirectoryName(item.Text);
                                nomeArquivo = Path.Combine(pastaArquivo, nomeArquivo) + ".xlsx";

                                converterHelper.JsonExcel(item.Text, nomeArquivo);
                            }

                            Tools.AbrirPastaSelecionandoArquivo(nomeArquivo);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private bool ValidarCampos()
        {
            if (listView1.Visible == true)
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                    if (!File.Exists(item.Text))
                    {
                        MessageBox.Show("Arquivo não existe:" + Environment.NewLine + item.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
            }

            else
            {
                if (comboBoxSistema.SelectedIndex == -1 || comboBoxSistema.SelectedIndex == -1 || string.IsNullOrWhiteSpace(txtEstabelecimentoID.Text))
                    return false;
            }

            return true;
        }


        void OcultarCampos()
        {
            foreach (Control control in this.Controls)
                control.Visible = false;

            label4.Visible = true;
            comboBoxImportacao.Visible = true;
            menuStrip1.Visible = true;

            txtEstabelecimentoID.Visible = true;
            lbEstabelecimento.Visible = true;
        }

        void AlterarNomesCampos()
        {
        }

        void MostrarCampos()
        {
            OcultarCampos();

            if (comboBoxImportacao.SelectedIndex > -1)
            {
                if (comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex] == "JSON")
                {
                    listView1.Visible = true;
                    btnAddToList.Visible = true;
                    btnDelFromList.Visible = true;
                    btnImportar.Visible = true;
                }

                else if (comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex] == "Todos")
                {
                    listView1.Visible = true;
                    btnAddToList.Visible = true;
                    btnDelFromList.Visible = true;
                    btnImportar.Visible = true;
                }

                else
                {
                    comboBoxSistema.Visible = true;
                    label3.Visible = true;

                    if (comboBoxSistema.SelectedIndex > -1 && comboBoxImportacao.SelectedIndex > -1)
                    {
                        AlterarNomesCampos();
                        btnImportar.Visible = true;
                    }
                }
            }
        }

        void NomeArquivoOpenFile()
        {
            nomeArquivoExcel = "";
        }

        private void comboBoxSistema_SelectedIndexChanged(object sender, EventArgs e)
        {
            MostrarCampos();
            NomeArquivoOpenFile();
        }

        private void comboBoxImportacao_SelectedIndexChanged(object sender, EventArgs e)
        {
            janelaArquivoExcel = comboBoxImportacao.Text;

            MostrarCampos();
            NomeArquivoOpenFile();
        }

        private void txtPessoaID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtEstabelecimentoID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtLoginID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void btnDelFromList_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.SelectedItems)
                listView1.Items.Remove(item);
        }

        private void btnAddToList_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            if (comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex] == "Todos")
                openFileDialog.Filter = "Arquivo Excel |*.csv";
            //openFileDialog.Filter = "Arquivo Excel |*.csv;*.xlsx";
            else if (comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex] == "JSON")
                openFileDialog.Filter = "Arquivo Json |*.json";

            openFileDialog.Title = "Selecione os arquivos";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                foreach (var file in openFileDialog.FileNames)
                    listView1.Items.Add(file);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var textoLinhas = Tools.LerConfig();

            try
            {
                Tools.salvarNaPasta = textoLinhas[0];
                Tools.ultimaPasta = textoLinhas[1];
                Tools.ultimoEstabelecimentoID = textoLinhas[2];
                Tools.ultimoEstabelecimento = textoLinhas[3];
                Tools.ultimoAntigoSistema = textoLinhas[4];
                Tools.ultimoinputDB = textoLinhas[5];
                Tools.ultimoinputDBContratos = textoLinhas[6];

                comboBoxSistema.SelectedIndex = int.Parse(Tools.ultimoAntigoSistema);
                txtEstabelecimentoID.Text = Tools.ultimoEstabelecimentoID;
            }
            catch
            {
                Tools.SalvarConfig();
            }

            if (!Directory.Exists(Tools.salvarNaPasta))
            {
                Tools.SalvarConfig();
                MessageBox.Show("Configure a pasta de saída clicando em \"Configurações\"", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void salvarNaPastaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pasta = AbrirPasta();

            if (!string.IsNullOrEmpty(pasta))
            {
                Tools.salvarNaPasta = pasta;
                Tools.SalvarConfig();
            }
        }

        private string AbrirPasta(string titulo = "Abrir")
        {
            string retorno = "";
            var folderBrowser = new OpenFileDialog();
            folderBrowser.ValidateNames = false;
            folderBrowser.InitialDirectory = Tools.salvarNaPasta;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            folderBrowser.Filter = "|Pasta";
            folderBrowser.FileName = "Abrir Pasta";
            folderBrowser.Title = titulo;
            if (folderBrowser.ShowDialog() == DialogResult.OK)
                retorno = Path.GetDirectoryName(folderBrowser.FileName);
            return retorno;
        }

        private void btnExcel2_Click_1(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                Filter = "Arquivo Excel |*.xlsx",
                Title = "Selecione um arquivo",
                InitialDirectory = Tools.ultimaPasta
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
                Tools.SalvarConfig();
            }
        }

        private void btnReferencia_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                Filter = "Arquivo Excel |*.xlsx",
                Title = "Selecione um arquivo",
                InitialDirectory = Tools.ultimaPasta
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
                Tools.SalvarConfig();
            }
        }

        private void abrirPastaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tools.AbrirPastaExplorer(Tools.salvarNaPasta);
        }

        private void btnPessoas_Click(object sender, EventArgs e)
        {
            var arquivoPessoas = EscolherArquivoExcel("Arquivo Pessoas.xlsx");

            if (string.IsNullOrEmpty(arquivoPessoas) == false)
                try
                {
                    var excelHelper = new ExcelHelper();
                    var workbook = excelHelper.LerExcel(arquivoPessoas);
                    var sheet = workbook.GetSheetAt(0);
                    excelHelper.InitializeDictionaryPessoas(sheet);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
        }

        private void btnRecebiveis_Click(object sender, EventArgs e)
        {
            var arquivoRecebiveis = EscolherArquivoExcel("Recebiveis");

            if (string.IsNullOrEmpty(arquivoRecebiveis) == false)
                try
                {
                    var excelHelper = new ExcelHelper();
                    var workbook = excelHelper.LerExcel(arquivoRecebiveis);
                    var sheet = workbook.GetSheetAt(0);
                    excelHelper.InitializeDictionaryRecebiveis(sheet);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
        }

        string EscolherArquivoExcel(string titulo = "Selecione um arquivo")
        {
            string retorno = "";

            var openFileDialog = new OpenFileDialog()
            {
                Filter = "Arquivo Excel |*.xlsx;*.csv",
                Title = titulo,
                InitialDirectory = Tools.ultimaPasta
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                retorno = openFileDialog.FileName;
                Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
                Tools.SalvarConfig();
            }

            return retorno;
        }

        private void OpenFormImportarDataBase(object sender, EventArgs e)
        {
            ImportacaoDataBase importacaoDataBase = new ImportacaoDataBase();
            importacaoDataBase.ShowDialog();
        }

        private void txtEstabelecimentoID_Leave(object sender, EventArgs e)
        {
            Tools.ultimoEstabelecimentoID = txtEstabelecimentoID.Text;
            Tools.SalvarConfig();
        }
    }
}
