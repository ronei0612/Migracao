using ExcelDataReader;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util.Collections;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using static System.Runtime.InteropServices.JavaScript.JSType;

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
					foreach (ListViewItem item in listView1.Items)
						Importar(item.Text);

					MessageBox.Show("<div class='msgResult iconOk icon-info-round icon-size2'>Migração<p class='msgSubResult'>Sucesso</p></div>", "Migração concluída", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
		}

		private bool ValidarCampos()
		{
			if (comboBoxSistema.SelectedIndex == -1 || comboBoxSistema.SelectedIndex == -1 || string.IsNullOrWhiteSpace(maskedTxtEstabelecimento.Text)
				 || string.IsNullOrWhiteSpace(textBoxExcel1.Text) || string.IsNullOrWhiteSpace(textBoxExcel2.Text))
				return false;

			if (!File.Exists(textBoxExcel1.Text))
			{
				MessageBox.Show("Arquivo não existe:" + Environment.NewLine + textBoxExcel1.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}
			else if (Path.GetExtension(textBoxExcel1.Text).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
			{
				MessageBox.Show("Arquivo não é um Excel (.xlsx):" + Environment.NewLine + textBoxExcel1.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}

			if (!File.Exists(textBoxExcel2.Text))
			{
				MessageBox.Show("Arquivo não existe:" + Environment.NewLine + textBoxExcel2.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}
			else if (Path.GetExtension(textBoxExcel2.Text).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
			{
				MessageBox.Show("Arquivo não é um Excel (.xlsx):" + Environment.NewLine + textBoxExcel2.Text, "Erro de Arquivo!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}

			return true;
		}

		private void Importar(string arquivoExcel)
		{
			IWorkbook workbook;
			var excelHelper = new ExcelHelper();
			DateTime dataMinima = new DateTime(1900, 01, 01), dataMaxima = new DateTime(2079, 06, 06), dataHoje = DateTime.Now;
			var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";

			var mascaraCPF = "000.000.000-00";
			mascaraCPF = mascaraCPF.Split('.')[0].Replace(".", @"\.").Replace("-", @"\-");
			var mascaraCPFLenth = Regex.Replace(mascaraCPF, "[^0-9]", "").Length.ToString();

			try
			{
				workbook = excelHelper.LerExcel(arquivoExcel);
			}
			catch (Exception ex)
			{
				throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
			}

			var cabecalhos = excelHelper.GetCabecalhosExcel(workbook);
			var linhas = excelHelper.GetLinhasExcel(workbook);

			try
			{
				var dados = new Dictionary<string, object[]>();

				var linhasCount = linhas.Count;

				var nomeCompleto = new string[linhasCount];
				var cpf = new string[linhasCount];
				var numcadastro = new int[linhasCount];
				var consumidorID = new int[linhasCount];
				var codigoAntigo = new int[linhasCount];
				var pessoaID = new int[linhasCount];

				foreach (var linha in linhas)
				{
					indiceLinha++;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim();
							tituloColuna = cabecalhos[celula.Address.Column];
							colunaLetra = celula.Address.ToString();

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								//if (!dados.ContainsKey(tituloColuna))
								//{
								//	dados[tituloColuna] = new List<object>();
								//}
								//dados[tituloColuna].Add(int.Parse(celulaValor));

								switch (tituloColuna)
								{
									case "numcadastro":
										numcadastro[indiceLinha - 2] = int.Parse(celulaValor);
										break;
									case "primeironome":
										nomeCompleto[indiceLinha - 2] = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "cpf":
										cpf[indiceLinha - 2] = celulaValor.Contains(".") && celulaValor.Contains("-") && celulaValor.Length <= 14 ? celulaValor : celulaValor.Length == int.Parse(mascaraCPFLenth) ? Convert.ToUInt64(celulaValor).ToString(mascaraCPF) : "";
										break;
								}
							}
						}
					}
				}

				dados.Add("numcadastro", numcadastro.Cast<object>().ToArray());
				dados.Add("nomeCompleto", nomeCompleto.Cast<object>().ToArray());
				dados.Add("cpf", cpf.Cast<object>().ToArray());

				//GravarExcel("asdf", dados);
				var insert = GerarSqlInsert("asdfff", dados);
				File.WriteAllText("aaaa.sql", insert);

			}

			catch (Exception error)
			{
				var mensagemErro = $"Falha na linha {indiceLinha}, coluna {colunaLetra}, Valor esperado: {tituloColuna}, valor da célula: \"{celulaValor}\": {error.Message}";

				if (!string.IsNullOrWhiteSpace(variaveisValor))
					mensagemErro += Environment.NewLine + "Variáveis" + Environment.NewLine + variaveisValor;

				throw new Exception(mensagemErro);
			}
		}

		private void GravarExcel(string nomeArquivo, Dictionary<string, object[]> linhas)
		{
			// Criando um novo arquivo Excel
			IWorkbook workbook = new XSSFWorkbook();
			ISheet sheet = workbook.CreateSheet("Dados");

			// Escrevendo cabeçalhos
			IRow headerRow = sheet.CreateRow(0);
			//for (int i = 0; i < cabecalhos.Count; i++)
			//{
			//	headerRow.CreateCell(i).SetCellValue(cabecalhos[i]);
			//}

			var cabecalhos = new List<string>(linhas.Keys);
			for (int i = 0; i < cabecalhos.Count; i++)
			{
				headerRow.CreateCell(i).SetCellValue(cabecalhos[i]);
			}

			// Escrevendo dados
			int rowIndex = 1;
			foreach (var linha in linhas)
			{
				IRow row = sheet.CreateRow(rowIndex++);
				for (int i = 0; i < linha.Value.Length; i++)
				{
					row.CreateCell(i).SetCellValue(linha.Value[i].ToString());
				}
			}

			// Salvando o arquivo
			using (FileStream stream = new FileStream(nomeArquivo + ".xlsx", FileMode.Create, FileAccess.Write))
			{
				workbook.Write(stream);
			}
		}

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

		public string GerarSqlInsert(string tableName, Dictionary<string, object[]> dataDict)
		{
			var sql = new StringBuilder($"INSERT INTO {tableName} (");

			// Adiciona os nomes das colunas
			foreach (var key in dataDict.Keys)
			{
				sql.Append($"{key}, ");
			}

			// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
			sql.Remove(sql.Length - 2, 2).Append(") VALUES ");

			// Adiciona os valores das colunas para cada linha
			for (int i = 0; i < dataDict.Values.First().Length; i++)
			{
				sql.Append('(');
				foreach (var valueArray in dataDict.Values)
				{
					sql.Append($"'{valueArray[i]}', ");
				}
				sql.Remove(sql.Length - 2, 2).Append("), ");
			}

			// Remove a última vírgula e espaço e adiciona um ponto e vírgula
			sql.Remove(sql.Length - 2, 2).Append(';');

			return sql.ToString();
		}

		void MostrarCamposExcel()
		{
			maskedTxtEstabelecimento.Focus();
			maskedTxtEstabelecimento.SelectAll();

			if (comboBoxSistema.SelectedIndex > -1 && comboBoxImportacao.SelectedIndex > -1 && !string.IsNullOrEmpty(maskedTxtEstabelecimento.Text))
			{
				textBoxExcel1.Visible = true;
				textBoxExcel2.Visible = true;
				btnExcel.Visible = true;
				btnExcel2.Visible = true;
				labelExcel1.Text = comboBoxImportacao.Text;
				labelExcel1.Visible = true;
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
	}
}
