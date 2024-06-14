using Migracao.Imports;
using Migracao.Models;
using Migracao.Sistems;
using Migracao.Utils;
using System.Windows.Forms;

namespace Migracao
{
	public partial class Form1 : Form
	{
		private const string arquivoConfig = "config.config";
		private string nomeArquivoExcel = "";
		private string janelaArquivoExcel = "Selecione um arquivo";

		public Form1()
		{
			InitializeComponent();
		}

		private void btnExcel_Click(object sender, EventArgs e)
		{
			using var openFileDialog = new OpenFileDialog
			{
				Filter = "Arquivo Excel |*.xlsx",
				Title = janelaArquivoExcel,
				FileName = nomeArquivoExcel,
				InitialDirectory = Tools.ultimaPasta
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				textBoxExcel1.Text = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				SalvarConfiguracoes();
			}
		}

		private void btnImportar_Click(object sender, EventArgs e)
		{
			Tools.ultimoEstabelecimentoID = txtEstabelecimentoID.Text;
			SalvarConfiguracoes();

			if (ValidarCampos())
			{
				try
				{
					if (listView1.Visible)
					{
						if (comboBoxImportacao.Text.Equals("todos", StringComparison.CurrentCultureIgnoreCase))
						{
							var odontoCompany = new OdontoCompany();
							odontoCompany.LerArquivos(txtEstabelecimentoID.Text, listView1);

							MessageBox.Show("Sucesso");
						}
						else if (comboBoxImportacao.Text.Equals("json", StringComparison.CurrentCultureIgnoreCase))
						{
							var converterHelper = new ConverterHelper();
							var nomeArquivo = "";
							foreach (ListViewItem item in listView1.Items)
							{
								nomeArquivo = Path.Combine(
									Path.GetDirectoryName(item.Text),
									Tools.TratarCaracteres(Path.GetFileNameWithoutExtension(item.Text)) + ".xlsx"
								);
								converterHelper.JsonExcel(item.Text, nomeArquivo);
							}

							Tools.AbrirPastaSelecionandoArquivo(nomeArquivo);
						}
					}
					else
					{
						var sistema = comboBoxSistema.Text.ToLower();
						var importacao = comboBoxImportacao.Text.ToLower();

						switch (sistema)
						{
							case "dentaloffice":
								var dentalOffice = new DentalOffice();
								switch (importacao)
								{
									case "pacientes":
										dentalOffice.ImportarPacientes(textBoxExcel1.Text, int.Parse(txtEstabelecimentoID.Text));
										break;
									case "recebidos":
										if (!string.IsNullOrEmpty(txtPessoaID.Text))
										{
											dentalOffice.ImportarRecebidos(
												textBoxExcel1.Text,
												txtReferencia.Text,
												int.Parse(txtEstabelecimentoID.Text),
												int.Parse(txtPessoaID.Text),
												int.Parse(txtLoginID.Text)
											);
										}
										break;
									case "pagos":
										if (!string.IsNullOrEmpty(txtPessoaID.Text))
										{
											dentalOffice.ImportarPagos(
												textBoxExcel1.Text,
												txtReferencia.Text,
												int.Parse(txtEstabelecimentoID.Text),
												int.Parse(txtPessoaID.Text),
												int.Parse(txtLoginID.Text)
											);
										}
										break;
								}
								break;
							case "odontocompany":
								var importacoes = new Importacoes();
								switch (importacao)
								{
									case "fornecedores":
										importacoes.ImportarFornecedores(
											textBoxExcel1.Text,
											txtReferencia.Text,
											int.Parse(txtEstabelecimentoID.Text),
											int.Parse(txtLoginID.Text)
										);
										break;
									case "pessoas":
										importacoes.ImportarPessoas(
											textBoxExcel1.Text,
											int.Parse(txtEstabelecimentoID.Text),
											int.Parse(txtLoginID.Text)
										);
										break;
									case "recebíveis":
										importacoes.ImportarRecebiveis(
											textBoxExcel1.Text,
											int.Parse(txtEstabelecimentoID.Text),
											int.Parse(txtPessoaID.Text),
											int.Parse(txtLoginID.Text)
										);
										break;
									case "preços":
										importacoes.ImportarPrecos(
											textBoxExcel1.Text,
											int.Parse(txtEstabelecimentoID.Text),
											int.Parse(txtLoginID.Text),
											txtReferencia.Text
										);
										break;
									case "agendamentos":
										importacoes.ImportarAgenda(
											textBoxExcel1.Text,
											int.Parse(txtEstabelecimentoID.Text),
											txtReferencia.Text,
											int.Parse(txtLoginID.Text),
											int.Parse(txtPessoaID.Text)
										);
										break;
								}
								break;
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
			if (listView1.Visible)
			{
				foreach (ListViewItem item in listView1.SelectedItems)
				{
					if (!File.Exists(item.Text))
					{
						MessageBox.Show("Arquivo não existe:" + Environment.NewLine + item.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}
				}
			}
			else
			{
				if (comboBoxSistema.SelectedIndex == -1 || comboBoxImportacao.SelectedIndex == -1 ||
					string.IsNullOrWhiteSpace(txtEstabelecimentoID.Text) || string.IsNullOrWhiteSpace(textBoxExcel1.Text) ||
					string.IsNullOrEmpty(txtLoginID.Text))
				{
					return false;
				}

				if (!File.Exists(textBoxExcel1.Text))
				{
					MessageBox.Show("Arquivo não existe:" + Environment.NewLine + textBoxExcel1.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return false;
				}
				else if (!Path.GetExtension(textBoxExcel1.Text).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
				{
					MessageBox.Show("Arquivo não é um Excel (.xlsx):" + Environment.NewLine + textBoxExcel1.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return false;
				}

				if (txtReferencia.Visible && !string.IsNullOrWhiteSpace(txtReferencia.Text))
				{
					if (!File.Exists(txtReferencia.Text))
					{
						MessageBox.Show("Arquivo não existe:" + Environment.NewLine + txtReferencia.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}
					else if (!Path.GetExtension(txtReferencia.Text).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
					{
						MessageBox.Show("Arquivo não é um Excel (.xlsx):" + Environment.NewLine + txtReferencia.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}
				}

				if (txtExcel2.Visible && !string.IsNullOrWhiteSpace(txtExcel2.Text))
				{
					if (!File.Exists(txtExcel2.Text))
					{
						MessageBox.Show("Arquivo não existe:" + Environment.NewLine + txtExcel2.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}
				}

				if (txtPessoaID.Visible && string.IsNullOrWhiteSpace(txtPessoaID.Text))
				{
					MessageBox.Show("Preencher campo Responsável Financeiro (PessoaID):", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return false;
				}
			}

			return true;
		}

		private void OcultarCampos()
		{
			foreach (Control control in this.Controls)
			{
				control.Visible = false;
			}

			label4.Visible = true;
			comboBoxImportacao.Visible = true;
			menuStrip1.Visible = true;

			label2.Visible = true;
			label6.Visible = true;
			txtPessoas.Visible = true;
			txtRecebiveis.Visible = true;
			btnPessoas.Visible = true;
			btnRecebiveis.Visible = true;

			txtEstabelecimentoID.Visible = true;
			lbEstabelecimento.Visible = true;
		}

		private void AlterarNomesCampos()
		{
			lbPessoaID.Text = "PessoaID RespFin:";

			var importacao = comboBoxImportacao.Text.ToLower();

			switch (importacao)
			{
				case "recebidos":
					lbExcel2.Text = "Recebíveis (Prod):";
					lbReferencia.Text = "Recebidos (Prod):";
					break;
				case "pacientes":
					lbReferencia.Text = "Pessoas (Prod):";
					break;
				case "agendamentos":
					lbReferencia.Text = "Agendamentos (Prod):";
					lbPessoaID.Text = "FuncionarioID Dent:";
					break;
			}
		}

		private void MostrarCampos()
		{
			OcultarCampos();

			if (comboBoxImportacao.SelectedIndex > -1)
			{
				var importacao = comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex].ToString();

				if (importacao == "JSON" || importacao == "Todos")
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
					label5.Visible = true;
					txtLoginID.Visible = true;

					if (comboBoxSistema.SelectedIndex > -1)
					{
						labelExcel1.Text = comboBoxImportacao.Text + ":";
						labelExcel1.Visible = true;
						textBoxExcel1.Visible = true;
						btnExcel.Visible = true;

						var importacaoLower = comboBoxImportacao.Text.ToLower();
						if (importacaoLower == "fornecedores" ||
							importacaoLower == "recebidos" ||
							importacaoLower == "preços procedimentos" ||
							importacaoLower == "tabela de preços" ||
							importacaoLower == "funcionarios" ||
							importacaoLower == "agendamentos")
						{
							lbReferencia.Visible = true;
							txtReferencia.Visible = true;
							btnReferencia.Visible = true;

							if (importacaoLower == "tabela de preços" || importacaoLower == "recebidos")
							{
								lbExcel2.Visible = true;
								txtExcel2.Visible = true;
								btnExcel2.Visible = true;
							}

							if (importacaoLower == "recebidos" || importacaoLower == "agendamentos")
							{
								lbPessoaID.Visible = true;
								txtPessoaID.Visible = true;
							}
						}

						if (importacaoLower == "recebíveis")
						{
							lbPessoaID.Visible = true;
							txtPessoaID.Visible = true;
						}

						AlterarNomesCampos();
						btnImportar.Visible = true;
					}
				}
			}
		}

		private void NomeArquivoOpenFile()
		{
			nomeArquivoExcel = "";

			var sistema = comboBoxSistema.Text.ToLower();
			var importacao = comboBoxImportacao.Text.ToLower();

			switch (sistema)
			{
				case "odontocompany":
					switch (importacao)
					{
						case "recebíveis":
							nomeArquivoExcel = "CRD111";
							break;
						case "funcionarios":
							nomeArquivoExcel = "CED006";
							break;
						case "pacientes":
							nomeArquivoExcel = "EMD101";
							break;
						case "pessoas":
							nomeArquivoExcel = "Pacient:EMD101 | Funcion:CED006";
							break;
						case "recebidos":
							nomeArquivoExcel = "BXD111";
							break;
						case "tabela de preços":
							nomeArquivoExcel = "CED001";
							break;
						case "agendamentos":
							nomeArquivoExcel = "AGENDA";
							break;
					}
					break;
			}
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
			{
				listView1.Items.Remove(item);
			}
		}

		private void btnAddToList_Click(object sender, EventArgs e)
		{
			using var openFileDialog = new OpenFileDialog
			{
				Filter = comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex].ToString() == "Todos" ?
					"Arquivo Excel |*.csv" : "Arquivo Json |*.json",
				Title = "Selecione os arquivos",
				Multiselect = true
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				foreach (var file in openFileDialog.FileNames)
				{
					listView1.Items.Add(file);
				}
			}
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			CarregarConfiguracoes();

			if (!Directory.Exists(Tools.salvarNaPasta))
			{
				MessageBox.Show("Configure a pasta de saída clicando em \"Configurações\"", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
		}

		private void salvarNaPastaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			var pasta = AbrirPasta();

			if (!string.IsNullOrEmpty(pasta))
			{
				Tools.salvarNaPasta = pasta;
				SalvarConfiguracoes();
			}
		}

		private string AbrirPasta(string titulo = "Abrir")
		{
			string retorno = "";
			using var folderBrowser = new OpenFileDialog
			{
				ValidateNames = false,
				InitialDirectory = Tools.salvarNaPasta,
				CheckFileExists = false,
				CheckPathExists = true,
				Filter = "|Pasta",
				FileName = "Abrir Pasta",
				Title = titulo
			};

			if (folderBrowser.ShowDialog() == DialogResult.OK)
			{
				retorno = Path.GetDirectoryName(folderBrowser.FileName);
			}

			return retorno;
		}

		private void btnExcel2_Click_1(object sender, EventArgs e)
		{
			using var openFileDialog = new OpenFileDialog
			{
				Filter = "Arquivo Excel |*.xlsx",
				Title = "Selecione um arquivo",
				InitialDirectory = Tools.ultimaPasta
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				txtExcel2.Text = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				SalvarConfiguracoes();
			}
		}

		private void btnReferencia_Click(object sender, EventArgs e)
		{
			using var openFileDialog = new OpenFileDialog
			{
				Filter = "Arquivo Excel |*.xlsx",
				Title = "Selecione um arquivo",
				InitialDirectory = Tools.ultimaPasta
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				txtReferencia.Text = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				SalvarConfiguracoes();
			}
		}

		private void abrirPastaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Tools.AbrirPastaExplorer(Tools.salvarNaPasta);
		}

		private void btnPessoas_Click(object sender, EventArgs e)
		{
			var arquivoPessoas = EscolherArquivoExcel("Arquivo Pessoas.xlsx");

			if (!string.IsNullOrEmpty(arquivoPessoas))
			{
				try
				{
					var excelHelper = new ExcelHelper();
					var workbook = excelHelper.LerExcel(arquivoPessoas);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionaryPessoas(sheet);

					txtPessoas.Text = arquivoPessoas;
				}
				catch (Exception ex)
				{
					txtPessoas.Text = "";
					MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
		}

		private void btnRecebiveis_Click(object sender, EventArgs e)
		{
			var arquivoRecebiveis = EscolherArquivoExcel("Recebiveis");

			if (!string.IsNullOrEmpty(arquivoRecebiveis))
			{
				try
				{
					var excelHelper = new ExcelHelper();
					var workbook = excelHelper.LerExcel(arquivoRecebiveis);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionaryRecebiveis(sheet);

					txtRecebiveis.Text = arquivoRecebiveis;
				}
				catch (Exception ex)
				{
					txtRecebiveis.Text = "";
					MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
		}

		private string EscolherArquivoExcel(string titulo = "Selecione um arquivo")
		{
			string retorno = "";

			using var openFileDialog = new OpenFileDialog
			{
				Filter = "Arquivo Excel |*.xlsx;*.csv",
				Title = titulo,
				InitialDirectory = Tools.ultimaPasta
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				retorno = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				SalvarConfiguracoes();
			}

			return retorno;
		}

		private void CarregarConfiguracoes()
		{
			if (!File.Exists(arquivoConfig))
			{
				SalvarConfiguracoes();
			}

			var textoLinhas = File.ReadAllLines(arquivoConfig);

			try
			{
				Tools.salvarNaPasta = textoLinhas[0];
				Tools.ultimaPasta = textoLinhas[1];
				Tools.ultimoEstabelecimentoID = textoLinhas[2];
			}
			catch
			{
				SalvarConfiguracoes();
			}
		}

		private void SalvarConfiguracoes()
		{
			File.WriteAllText(arquivoConfig,
				Tools.salvarNaPasta + Environment.NewLine +
				Tools.ultimaPasta + Environment.NewLine +
				Tools.ultimoEstabelecimentoID
			);
		}
	}
}