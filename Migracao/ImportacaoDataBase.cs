﻿using Migracao.Models.Interfaces;
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
        private string _sistemaOrigemIndex;
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
            if (btnImportar.Text.Contains("Executar"))
            {
                btnImportar.Enabled = false;

                backgroundThreadStart = new ThreadStart(ExecutarImportacao);
                backgroundThread = new Thread(backgroundThreadStart);
                backgroundThread.Start();
            }
        }

        private void ExecutarImportacao()
        {
            this.Invoke((MethodInvoker)delegate {
                _sistemaOrigem = comboBoxSistema.SelectedItem.ToString();
                _sistemaOrigemIndex = comboBoxSistema.SelectedIndex.ToString();
                _tabela = comboTabelas.SelectedItem?.ToString();
                _pathDB = inputDB.Text;
                _pathDBContratos = inputDBContratos.Text;
                _dataBaseName = inputDataBaseName.Text.ToString();
            });

            Tools.ultimoAntigoSistema = _sistemaOrigemIndex;
            Tools.ultimoinputDB = _pathDB;
            Tools.ultimoinputDBContratos = _pathDBContratos;
            Tools.SalvarConfig();

            if (string.IsNullOrEmpty(_tabela))
                MessageBox.Show("Para continuar, selecione uma das opções de tabela para importação!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (string.IsNullOrEmpty(_pathDB) || (_sistemaOrigem == "OdontoCompany" && string.IsNullOrEmpty(_pathDBContratos)))
                MessageBox.Show("Por favor, valide o caminho das DBs desejadas", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

            try
            {
                //Identifica a classe, baseado na escolha de sistema, cria a instancia e chama o método através dela
                GetImportacaoMetodo();

                this.Invoke((MethodInvoker)delegate
                {
                    btnImportar.Enabled = true;
                });

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
                    // Principais
                    case "TUDO":
                        instance.DataBaseImportacaoDevClinico();
                        instance.DataBaseImportacaoManutencoes();
                        instance.DataBaseImportacaoPacientesDentistas();
                        instance.DataBaseImportacaoProcedimentosPrecos();
                        instance.DataBaseImportacaoPagosExigiveis();
                        instance.DataBaseImportacaoProcedimentos();
                        instance.DataBaseImportacaoFinanceiroRecebiveis();
                        instance.DataBaseImportacaoAgendamentos();
                        instance.DataBaseImportacaoDentistas();
                        instance.DataBaseImportacaoRecebiveisHistVenda();
                        break;
                    case "Agendamentos/DesenvClínico":
                        instance.DataBaseImportacaoDevClinico();
                        break;
                    case "Procedimentos/Manutenções":
                        instance.DataBaseImportacaoManutencoes();
                        break;
                    case "Pacientes/Dentistas":
                        instance.DataBaseImportacaoPacientesDentistas();
                        break;
                    case "Procedimentos Tabela Preços":
                        instance.DataBaseImportacaoProcedimentosPrecos();
                        break;
                    case "Recebíveis/Pagos/Exigíveis":
                        instance.DataBaseImportacaoPagosExigiveis();
                        break;

                        // Complementares
                        //case "Procedimentos":
                        //    instance.DataBaseImportacaoProcedimentos();
                        //    break;

                        //case "Financeiro (Recebíveis)":
                        //instance.DataBaseImportacaoFinanceiroRecebiveis();
                        //    break;
                        //case "Agendamentos":
                        //    instance.DataBaseImportacaoAgendamentos();
                        //    break;
                        //case "Dentistas":
                        //    instance.DataBaseImportacaoDentistas();
                        //    break;
                        //case "Recebíveis Histórico Vendas":
                        //    instance.DataBaseImportacaoRecebiveisHistVenda();
                        //    break;                    
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
