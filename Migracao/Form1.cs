using Migracao.Sistems;
using Migracao.Utils;

namespace Migracao
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

						else if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase))
							if (!string.IsNullOrEmpty(txtPessoaID.Text))
								dentalOffice.ImportarRecebidos(textBoxExcel1.Text, textBoxExcel2.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtPessoaID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("pagos", StringComparison.CurrentCultureIgnoreCase))
							if (!string.IsNullOrEmpty(txtPessoaID.Text))
								dentalOffice.ImportarPagos(textBoxExcel1.Text, textBoxExcel2.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtPessoaID.Text), int.Parse(txtLoginID.Text));
					}

					else if (comboBoxSistema.Text.Equals("odontocompany", StringComparison.CurrentCultureIgnoreCase))
					{
						var odontoCompany = new OdontoCompany();

						if (comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarPacientes(textBoxExcel1.Text, textBoxExcel2.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtLoginID.Text));

						else if (comboBoxImportacao.Text.Equals("fornecedores", StringComparison.CurrentCultureIgnoreCase))
							odontoCompany.ImportarFornecedores(textBoxExcel1.Text, textBoxExcel2.Text, int.Parse(txtEstabelecimentoID.Text), int.Parse(txtLoginID.Text));
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

				if (textBoxExcel2.Visible == true && !string.IsNullOrWhiteSpace(textBoxExcel2.Text))
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
			}

			return true;
		}

		void OcultarElementos()
		{
			foreach (Control control in this.Controls)
				control.Visible = false;

			label4.Visible = true;
			comboBoxImportacao.Visible = true;
		}

		void MostrarCampos()
		{
			OcultarElementos();

			if (comboBoxImportacao.SelectedIndex > -1)
			{
				if (comboBoxImportacao.Items[comboBoxImportacao.SelectedIndex] == "JSON")
				{
					listView1.Visible = true;
					btnAddToList.Visible = true;
					btnDelFromList.Visible = true;
					btnImportar.Visible = true;
				}

				else {
					comboBoxSistema.Visible = true;
					txtEstabelecimentoID.Visible = true;
					label1.Visible = true;
					label3.Visible = true;
					label5.Visible = true;
					txtLoginID.Visible = true;

					if (comboBoxSistema.SelectedIndex > -1 && comboBoxImportacao.SelectedIndex > -1)
					{
						labelExcel1.Text = comboBoxImportacao.Text;
						labelExcel1.Visible = true;
						textBoxExcel1.Visible = true;
						btnExcel.Visible = true;

						if (comboBoxImportacao.Text.Equals("recebidos", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("pacientes", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("fornecedores", StringComparison.CurrentCultureIgnoreCase)
							|| comboBoxImportacao.Text.Equals("pagos", StringComparison.CurrentCultureIgnoreCase))
						{
							labelExcel2.Text = "Referência";
							labelExcel2.Visible = true;
							textBoxExcel2.Visible = true;
							btnExcel2.Visible = true;
							label2.Visible = true;
							txtPessoaID.Visible = true;
						}

						btnImportar.Visible = true;
					}
				}
			}
		}

		private void comboBoxSistema_SelectedIndexChanged(object sender, EventArgs e)
		{
			MostrarCampos();
		}

		private void comboBoxImportacao_SelectedIndexChanged(object sender, EventArgs e)
		{
			MostrarCampos();
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
	}
}
