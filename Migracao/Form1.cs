using Migracao.Sistems;
using Migracao.Utils;

namespace Migracao
{
	public partial class Form1 : Form
	{
		string arquivoConfig = "config.config";
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
				textBoxExcel1.Text = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				File.WriteAllText(arquivoConfig, Tools.salvarNaPasta + Environment.NewLine + Tools.ultimaPasta);
			}
		}

		private void btnImportar_Click(object sender, EventArgs e)
		{
			if (ValidarCampos())
			{
				try
				{
					if (listView1.Visible == true)
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

					else if (comboBoxSistema.Text.Equals("dentaloffice", StringComparison.CurrentCultureIgnoreCase))
					{
						var dentalOffice = new DentalOffice();

						if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase))
							dentalOffice.ImportarPacientes(textBoxExcel1.Text, int.Parse(txtEstabelecimentoID.Text));

						else if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase)
							&& !string.IsNullOrEmpty(txtPessoaID.Text))
							dentalOffice.ImportarRecebidos(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtPessoaID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("pagos", StringComparison.CurrentCultureIgnoreCase)
							&& !string.IsNullOrEmpty(txtPessoaID.Text))
							dentalOffice.ImportarPagos(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtPessoaID.Text), int.Parse(txtLoginID.Text));
					}

					else if (comboBoxSistema.Text.Equals("odontocompany", StringComparison.CurrentCultureIgnoreCase))
					{
						var odontoCompany = new OdontoCompany();

						if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarPacientes(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("fornecedores", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarFornecedores(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("funcionarios", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarFornecedores(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("pessoas", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarFornecedores(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("recebíveis", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarRecebiveis(textBoxExcel1.Text, txtReferencia.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtPessoaID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarRecebidos(textBoxExcel1.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtPessoaID.Text), int.Parse(txtLoginID.Text), txtExcel2.Text, txtReferencia.Text);

						else if (comboBoxImportacao.Text.Equals("tabela de preços", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarPrecos(textBoxExcel1.Text, int.Parse(txtEstabelecimentoID.Text), txtReferencia.Text);
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
				if (comboBoxSistema.SelectedIndex == -1 || comboBoxSistema.SelectedIndex == -1 || string.IsNullOrWhiteSpace(txtEstabelecimentoID.Text)
					 || string.IsNullOrWhiteSpace(textBoxExcel1.Text) || string.IsNullOrEmpty(txtLoginID.Text))
					return false;

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

				if (txtReferencia.Visible == true && !string.IsNullOrWhiteSpace(txtReferencia.Text))
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

				if (txtExcel2.Visible == true && !string.IsNullOrWhiteSpace(txtExcel2.Text))
					if (!File.Exists(txtExcel2.Text))
					{
						MessageBox.Show("Arquivo não existe:" + Environment.NewLine + txtExcel2.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}

				if (txtPessoaID.Visible == true && string.IsNullOrWhiteSpace(txtPessoaID.Text))
				{
					MessageBox.Show("Preencher campo Responsável Financeiro (PessoaID):", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return false;
				}
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
		}

		void AlterarNomesCampos()
		{
			if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase))
			{
				lbExcel2.Text = "Recebíveis (Prod):";
				lbReferencia.Text = "Form Pgto (CRD013):";
			}
			else if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase) || comboBoxImportacao.Text.Equals("recebíveis", StringComparison.CurrentCultureIgnoreCase))
				lbReferencia.Text = "Pessoas (Prod):";
							//|| comboBoxImportacao.Text.Equals("fornecedores", StringComparison.CurrentCultureIgnoreCase)
							//|| comboBoxImportacao.Text.Equals("pagos", StringComparison.CurrentCultureIgnoreCase)
							
							//|| comboBoxImportacao.Text.Equals("tabela de preços", StringComparison.CurrentCultureIgnoreCase))
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

				else
				{
					comboBoxSistema.Visible = true;
					txtEstabelecimentoID.Visible = true;
					label1.Visible = true;
					label3.Visible = true;
					label5.Visible = true;
					txtLoginID.Visible = true;

					if (comboBoxSistema.SelectedIndex > -1 && comboBoxImportacao.SelectedIndex > -1)
					{
						labelExcel1.Text = comboBoxImportacao.Text + ":";
						labelExcel1.Visible = true;
						textBoxExcel1.Visible = true;
						btnExcel.Visible = true;

						if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("fornecedores", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("pessoas", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("pagos", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("recebíveis", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("tabela de preços", StringComparison.CurrentCultureIgnoreCase))
						{
							lbReferencia.Visible = true;
							txtReferencia.Visible = true;
							btnReferencia.Visible = true;

							if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase)
								|| comboBoxImportacao.Text.Equals("tabela de preços", StringComparison.CurrentCultureIgnoreCase))
							//if (comboBoxImportacao.Text.Equals("tabela de preços", StringComparison.CurrentCultureIgnoreCase))
							{
								lbExcel2.Visible = true;
								txtExcel2.Visible = true;
								btnExcel2.Visible = true;
							}
							
							if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase) || comboBoxImportacao.Text.Equals("recebíveis", StringComparison.CurrentCultureIgnoreCase))
							{
								lbPessoaID.Visible = true;
								txtPessoaID.Visible = true;
							}
						}

						AlterarNomesCampos();
						btnImportar.Visible = true;
					}
				}
			}
		}

		void NomeArquivoOpenFile()
		{
			nomeArquivoExcel = "";

			if (comboBoxSistema.Text.Equals("odontocompany", StringComparison.CurrentCultureIgnoreCase))
			{
				if (comboBoxImportacao.Text.Equals("recebíveis", StringComparison.CurrentCultureIgnoreCase))
					nomeArquivoExcel = "CRD111";
				else if (comboBoxImportacao.Text.Equals("funcionarios", StringComparison.CurrentCultureIgnoreCase))
					nomeArquivoExcel = "CED006";
				else if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase)
					nomeArquivoExcel = "EMD101";
				else if (comboBoxImportacao.Text.Equals("pessoas", StringComparison.CurrentCultureIgnoreCase))
					nomeArquivoExcel = "Pacient:EMD101 | Funcion:CED006";
				else if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase))
					nomeArquivoExcel = "BXD111";
				else if (comboBoxImportacao.Text.Equals("tabela de preços", StringComparison.CurrentCultureIgnoreCase))
					nomeArquivoExcel = "CED001";
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

		private void txtEstabelecimentoID_TextChanged(object sender, EventArgs e)
		{
		}

		private void txtPessoaID_TextChanged(object sender, EventArgs e)
		{
			//MostrarCampos();
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
			openFileDialog.Filter = "Arquivo Json |*.json";
			openFileDialog.Title = "Selecione um arquivo";

			if (openFileDialog.ShowDialog() == DialogResult.OK)
				listView1.Items.Add(openFileDialog.FileName);
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			if (!File.Exists(arquivoConfig))
				File.WriteAllText(arquivoConfig, Tools.salvarNaPasta + Environment.NewLine + Tools.ultimaPasta);

			var textoLinhas = File.ReadAllLines(arquivoConfig);

			Tools.salvarNaPasta = textoLinhas[0];
			Tools.ultimaPasta = textoLinhas[1];
		}

		private void salvarNaPastaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			string pasta = AbrirPasta();

			if (!string.IsNullOrEmpty(pasta))
			{
				Tools.salvarNaPasta = pasta;
				File.WriteAllText(arquivoConfig, Tools.salvarNaPasta + Environment.NewLine + Tools.ultimaPasta);
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
				txtExcel2.Text = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				File.WriteAllText(arquivoConfig, Tools.salvarNaPasta + Environment.NewLine + Tools.ultimaPasta);
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
				txtReferencia.Text = openFileDialog.FileName;
				Tools.ultimaPasta = Path.GetDirectoryName(openFileDialog.FileName);
				File.WriteAllText(arquivoConfig, Tools.salvarNaPasta + Environment.NewLine + Tools.ultimaPasta);
			}
		}

		private void abrirPastaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Tools.AbrirPastaExplorer(Tools.salvarNaPasta);
		}
	}
}
