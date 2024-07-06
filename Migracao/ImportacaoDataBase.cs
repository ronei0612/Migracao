using Migracao.Models.Interfaces;
using Migracao.Utils;

namespace Migracao
{
    public partial class ImportacaoDataBase : Form
    {
        string arquivoConfig = "config.config";
        string nomeArquivoExcel = "";
        string janelaArquivoExcel = "Selecione um arquivo";

        private string _pathDB;
        private string _pathDBContratos;
        private string _dataBaseName;
        private string _tabela;
        private string _sistemaOrigem;
        ThreadStart backgroundThreadStart;
        Thread backgroundThread;


        public ImportacaoDataBase()
        {
            InitializeComponent();
        }

        private void AntigoSistemaChanged(object sender, EventArgs e)
        {
            if (comboBoxSistema.Items[comboBoxSistema.SelectedIndex] == "DentalOffice")
            {
                panelDB.Visible = false;
                panelDBContratos.Visible = false;

                inputDBContratos.Text = string.Empty;

                panelDataBase.Visible = true;
            }

            if (comboBoxSistema.Items[comboBoxSistema.SelectedIndex] == "OdontoCompany")
            {
                panelDB.Visible = true;
                panelDB.Visible = true;
                panelDBContratos.Visible = true;

                inputDataBaseName.Text = string.Empty;

                panelDataBase.Visible = false;
            }
        }

        private void BtnPathDB_Click(object sender, EventArgs e)
        {
            SelecionarArquivo(inputDB);
        }

        private void BtnPathDBContratos_Click(object sender, EventArgs e)
        {
            SelecionarArquivo(inputDBContratos);
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            if (btnImportar.Text.Contains("Executar")) {
                btnImportar.Enabled = false;

                backgroundThread = new Thread(backgroundThreadStart);
                backgroundThread.Start();
            }
        }

        private void ExecutarImportacao()
        {
            _sistemaOrigem = comboBoxSistema.SelectedItem.ToString();

            _pathDB = inputDB.Text;
            _pathDBContratos = inputDBContratos.Text;

            Tools.ultimoAntigoSistema = comboBoxSistema.SelectedIndex.ToString();
            Tools.ultimoinputDB = _pathDB;
            Tools.ultimoinputDBContratos = _pathDBContratos;
            Tools.SalvarConfig();

            _dataBaseName = inputDataBaseName.Text.ToString();

            _tabela = comboTabelas.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(_tabela))
                MessageBox.Show("Para continuar, selecione uma das opções de tabela para importação!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (string.IsNullOrEmpty(_pathDB) || (_sistemaOrigem == "OdontoCompany" && string.IsNullOrEmpty(_pathDBContratos)))
                MessageBox.Show("Por favor, valide o caminho das DBs desejadas", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

            try
            {
                //Identifica a classe, baseado na escolha de sistema, cria a instancia e chama o método através dela
                GetImportacaoMetodo();

                btnImportar.Enabled = true;
                MessageBox.Show("Sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetImportacaoMetodo()
        {
            Type type = Type.GetType("Migracao.Sistems." + _sistemaOrigem);

            if (type != null && typeof(IDataBaseMigracao).IsAssignableFrom(type))
            {
                IDataBaseMigracao instance = (IDataBaseMigracao)Activator.CreateInstance(type, _dataBaseName, _pathDB, _pathDBContratos);

                switch (_tabela)
                {
                    case "TUDO":
                        instance.DataBaseImportacaoProcedimentos();
                        instance.DataBaseImportacaoDevClinico();
                        instance.DataBaseImportacaoManutencoes();
                        instance.DataBaseImportacaoPacientesDentistas();
                        instance.DataBaseImportacaoFinanceiroRecebiveis();
                        instance.DataBaseImportacaoAgendamentos();
                        instance.DataBaseImportacaoDentistas();
                        instance.DataBaseImportacaoRecebiveisHistVenda();
                        instance.DataBaseImportacaoProcedimentosPrecos();
                        instance.DataBaseImportacaoPagosExigiveis();
                        break;
                    case "Procedimentos":
                        instance.DataBaseImportacaoProcedimentos();
                        break;
                    case "Desenvolvimento Clínico":
                        instance.DataBaseImportacaoDevClinico();
                        break;
                    case "Manutenções":
                        instance.DataBaseImportacaoManutencoes();
                        break;
                    case "Pacientes/Dentistas":
                        instance.DataBaseImportacaoPacientesDentistas();
                        break;
                    case "Financeiro (Recebíveis)":
                        //    instance.DataBaseImportacaoFinanceiroRecebidos();

                        if (_tabela == "Financeiro (Recebíveis)")
                            instance.DataBaseImportacaoFinanceiroRecebiveis();
                        break;
                    case "Agendamentos":
                        instance.DataBaseImportacaoAgendamentos();
                        break;
                    case "Dentistas":
                        instance.DataBaseImportacaoDentistas();
                        break;
                    case "Recebíveis Histórico Vendas":
                        instance.DataBaseImportacaoRecebiveisHistVenda();
                        break;
                    case "Procedimentos Preços":
                        instance.DataBaseImportacaoProcedimentosPrecos();
                        break;
                    case "Recebíveis Pagos e Exigíveis":
                        instance.DataBaseImportacaoPagosExigiveis();
                        break;
                }
            }
            else
            {
                MessageBox.Show("Classe não encontrada ou não implementa a Interface!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelecionarArquivo(TextBox textBox)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Selecione um arquivo",
                InitialDirectory = Tools.ultimaPasta
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = openFileDialog.FileName;
                Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
                Tools.SalvarConfig();
            }
        }

        private void ImportacaoDataBase_Load(object sender, EventArgs e)
        {
            comboBoxSistema.SelectedIndex = int.Parse(Tools.ultimoAntigoSistema);
            inputDB.Text = Tools.ultimoinputDB;
            inputDBContratos.Text = Tools.ultimoinputDBContratos;
            comboTabelas.SelectedIndex = 0;

            backgroundThreadStart = new ThreadStart(ExecutarImportacao);
        }

        private void ImportacaoDataBase_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Hide();
        }

        private void BtnPathDB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Hide();
        }
    }
}
