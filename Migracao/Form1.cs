using Migracao.Utils;

namespace Migracao
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
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

            else if (string.IsNullOrWhiteSpace(txtEstabelecimentoID.Text))
                return false;

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

                else if (comboBoxImportacao.SelectedIndex > -1)
                    btnImportar.Visible = true;
            }
        }

        private void comboBoxImportacao_SelectedIndexChanged(object sender, EventArgs e)
        {
            MostrarCampos();
        }

        private void txtEstabelecimentoID_KeyPress(object sender, KeyPressEventArgs e)
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

            if (comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex] == "JSON")
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

        private void abrirPastaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tools.AbrirPastaExplorer(Tools.salvarNaPasta);
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
