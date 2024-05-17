using System.Diagnostics;

namespace Migração
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void btnExcel_Click(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			//openFileDialog.Filter = "Arquivo Excel |*.xls;*.xlsx";
			openFileDialog.Filter = "Arquivo Excel |*.xlsx";
			openFileDialog.Title = "Selecione um arquivo";

			if (openFileDialog.ShowDialog() == DialogResult.OK)
				textBoxExcel1.Text = openFileDialog.FileName;
			//ListViewItem item = new ListViewItem(filePath);
			//listView1.Items.Add(item);
		}

		private void btnExcel2_Click(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Arquivo Excel |*.xlsx";
			openFileDialog.Title = "Selecione um arquivo";

			if (openFileDialog.ShowDialog() == DialogResult.OK)
				textBoxExcel2.Text = openFileDialog.FileName;
		}

		private void btnDelExcel_Click(object sender, EventArgs e)
		{
			if (listView1.SelectedItems.Count > 0)
				listView1.Items.Remove(listView1.SelectedItems[0]);
			else
				MessageBox.Show("Por favor, selecione um item para remover.", "Nenhum item selecionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}

		private void btnImportar_Click(object sender, EventArgs e)
		{
			if (ValidarCampos())
			{
				try
				{
					//foreach (ListViewItem item in listView1.Items)
					//	Importar(item.Text);

					var pasta = Environment.ExpandEnvironmentVariables("%userprofile%\\Desktop");
					var arquivo = "Migracao_" + txtEstabelecimentoID.Text + "_DentalOffice_Pacientes";
					string caminhoDoArquivo = Path.Combine(pasta, arquivo);

					if (comboBoxSistema.Text.Equals("dentaloffice", StringComparison.CurrentCultureIgnoreCase))
					{
						if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase))
						{
							var dentalOffice = new DentalOffice();
							dentalOffice.ImportarPacientes(textBoxExcel1.Text, txtEstabelecimentoID.Text, caminhoDoArquivo);
						}

						else if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase))
						{
							var dentalOffice = new DentalOffice();
							dentalOffice.ImportarRecebidos(textBoxExcel1.Text, textBoxExcel2.Text, txtEstabelecimentoID.Text, txtPessoaID.Text, caminhoDoArquivo);
						}
					}
					else if (comboBoxSistema.Text.Equals("odontocompany"))
					{
						if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase))
						{
							var dentalOffice = new DentalOffice();
							dentalOffice.ImportarPacientes(textBoxExcel1.Text, txtEstabelecimentoID.Text, caminhoDoArquivo);
						}
					}

					// Application.StartupPath
					caminhoDoArquivo += ".xlsx";
					string argumento = "/select, \"" + caminhoDoArquivo + "\"";

					Process.Start("explorer.exe", argumento);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
		}

		private bool ValidarCampos()
		{
			if (comboBoxSistema.SelectedIndex == -1 || comboBoxSistema.SelectedIndex == -1 || string.IsNullOrWhiteSpace(txtEstabelecimentoID.Text)
				 || string.IsNullOrWhiteSpace(textBoxExcel1.Text))
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

			if (textBoxExcel2.Visible == true)
				if (!File.Exists(textBoxExcel2.Text))
				{
					MessageBox.Show("Arquivo não existe:" + Environment.NewLine + textBoxExcel2.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return false;
				}
				else if (!Path.GetExtension(textBoxExcel2.Text).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
				{
					MessageBox.Show("Arquivo não é um Excel (.xlsx):" + Environment.NewLine + textBoxExcel2.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return false;
				}

			return true;
		}

		//private void Importar(string arquivoExcel1, string arquivoExcel2, string sistema, string importacao)
		//{
		//	IWorkbook workbook;
		//	var excelHelper = new ExcelHelper();
		//	var dentalOffice = new DentalOffice();

		//	try
		//	{
		//		workbook = excelHelper.LerExcel(arquivoExcel1);
		//	}
		//	catch (Exception ex)
		//	{
		//		throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
		//	}

		//	var cabecalhos = excelHelper.GetCabecalhosExcel(workbook);
		//	var linhas = excelHelper.GetLinhasExcel(workbook);

		//	try
		//	{
		//		var dados = dentalOffice.ImportarPacientes(linhas, cabecalhos);				

		//		GravarExcel("asdf", dados);
		//		var insert = GerarSqlInsert("asdfff", dados);
		//		File.WriteAllText("aaaa.sql", insert);
		//	}

		//	catch (Exception error)
		//	{
		//		throw new Exception(error.Message);
		//	}
		//}



		//private IWorkbook LerExcel(Stream fileStream)
		//{
		//	// Determine the Excel format and create appropriate workbook instance
		//	if (Path.GetExtension(FileName).Equals(".xls", StringComparison.OrdinalIgnoreCase))
		//	{
		//		return new HSSFWorkbook(fileStream);
		//	}
		//	else if (Path.GetExtension(FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
		//	{
		//		return new XSSFWorkbook(fileStream);
		//	}
		//	else
		//	{
		//		throw new Exception("Formato de arquivo Excel não suportado.");
		//	}
		//}

		void MostrarCamposExcel()
		{
			txtEstabelecimentoID.Focus();
			txtEstabelecimentoID.SelectAll();

			if (comboBoxSistema.SelectedIndex > -1 && comboBoxImportacao.SelectedIndex > -1 && !string.IsNullOrEmpty(txtEstabelecimentoID.Text))
			{
				labelExcel1.Text = comboBoxImportacao.Text;
				labelExcel1.Visible = true;
				textBoxExcel1.Visible = true;
				btnExcel.Visible = true;

				if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase)
					|| comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase))
				{
					labelExcel2.Text = "Pacientes Referência";
					labelExcel2.Visible = true;
					textBoxExcel2.Visible = true;
					btnExcel2.Visible = true;
					label2.Visible = true;
					txtPessoaID.Visible = true;
				}
			}
			else
			{
				textBoxExcel1.Visible = false;
				textBoxExcel2.Visible = false;
				btnExcel.Visible = false;
				btnExcel2.Visible = false;
			}
		}

		private void comboBoxSistema_SelectedIndexChanged(object sender, EventArgs e)
		{
			MostrarCamposExcel();
		}

		private void comboBoxImportacao_SelectedIndexChanged(object sender, EventArgs e)
		{
			MostrarCamposExcel();
		}

		private void maskedTxtEstabelecimento_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
		{

		}

		private void maskedTxtEstabelecimento_TextChanged(object sender, EventArgs e)
		{
			MostrarCamposExcel();
		}

		private void maskedTxtEstabelecimento_Enter(object sender, EventArgs e)
		{
			maskedTxtEstabelecimento.SelectAll();
		}

		private void maskedTxtEstabelecimento_Click(object sender, EventArgs e)
		{
			maskedTxtEstabelecimento.SelectAll();
		}

		private void maskedTextBox2_Enter(object sender, EventArgs e)
		{
			maskedTextBox2.SelectAll();
		}

		private void maskedTextBox2_TextChanged(object sender, EventArgs e)
		{

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
			MostrarCamposExcel();
		}

		private void txtPessoaID_TextChanged(object sender, EventArgs e)
		{
			//MostrarCamposExcel();
		}
	}
}
