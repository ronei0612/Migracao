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
			var openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Arquivo Excel |*.xlsx";
			openFileDialog.Title = "Selecione um arquivo";

			if (openFileDialog.ShowDialog() == DialogResult.OK)
				textBoxExcel1.Text = openFileDialog.FileName;
		}

		private void btnExcel2_Click(object sender, EventArgs e)
		{
			var openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Arquivo Excel |*.xlsx";
			openFileDialog.Title = "Selecione um arquivo";

			if (openFileDialog.ShowDialog() == DialogResult.OK)
				textBoxExcel2.Text = openFileDialog.FileName;
		}

		private void btnImportar_Click(object sender, EventArgs e)
		{
			if (ValidarCampos())
			{
				try
				{
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
