using Migracao.Models;
using Migracao.Models.DentalOffice;
using Migracao.Models.Interfaces;
using Migracao.Sistems;
using Migracao.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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


        public ImportacaoDataBase()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

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

        private void ExecutarImportacao(object sender, EventArgs e)
        {
            _sistemaOrigem = comboBoxSistema.SelectedItem.ToString();

            _pathDB = inputDB.Text;
            _pathDBContratos = inputDBContratos.Text;

            _dataBaseName = inputDataBaseName.Text.ToString();

            _tabela = comboTabelas.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(_tabela))
                MessageBox.Show("Para continuar, selecione uma das opções de tabela para importação!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (string.IsNullOrEmpty(_pathDB) || (_sistemaOrigem == "OdontoCompany" && string.IsNullOrEmpty(_pathDBContratos)))
                MessageBox.Show("Por favor, valide o caminho das DBs desejadas", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //Identifica a classe, baseado na escolha de sistema, cria a instancia e chama o método através dela
            GetImportacaoMetodo();
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
                    case "AgendamentosDesenvClínico":
                        instance.DataBaseImportacaoDevClinico();
                        break;
                    case "ProcedimentosManutenções":
                        instance.DataBaseImportacaoManutencoes();
                        break;
                    case "PacientesDentistas":
                        instance.DataBaseImportacaoPacientesDentistas();
                        break;

                    case "ProcedimentosTabela Preços":
                        instance.DataBaseImportacaoProcedimentosPrecos();
                        break;
                    case "Recebíveis Pagos e Exigíveis":
                        instance.DataBaseImportacaoPagosExigiveis();
                        break;

                    // Complementares
                    case "Procedimentos":
                        instance.DataBaseImportacaoProcedimentos();
                        break;

                    case "Financeiro (Recebíveis)":
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
                File.WriteAllText(arquivoConfig, Tools.salvarNaPasta + Environment.NewLine + Tools.ultimaPasta + Environment.NewLine + Tools.ultimoEstabelecimentoID);
            }
        }
    }
}
