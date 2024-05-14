using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util.Collections;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

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
			{
				string filePath = openFileDialog.FileName;

				ListViewItem item = new ListViewItem(filePath);
				listView1.Items.Add(item);
			}
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
			try
			{				

				foreach (ListViewItem item in listView1.Items)
				{
					Importar(item.Text);
				}

				MessageBox.Show("<div class='msgResult iconOk icon-info-round icon-size2'>Migração<p class='msgSubResult'>Sucesso</p></div>", "Migração concluída", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
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
				var dados = new Dictionary<string, List<object>>();

				//var nomeCompleto = new List<string>();
				//var cpf = new List<string>(); 
				//var controle = new List<int>();
				//var consumidorID = new List<int>();
				//var codigoAntigo = new List<int>();
				//var pessoaID = new List<int>();

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
								if (!dados.ContainsKey(tituloColuna))
								{
									dados[tituloColuna] = new List<object>();
								}

								switch (tituloColuna)
								{
									case "ID":
									case "Controle":
									case "CodigoAntigo":
									case "PessoaID":
										dados[tituloColuna].Add(int.Parse(celulaValor));
										break;
									case "NomeCompleto":
										dados[tituloColuna].Add(celulaValor.Substring(0, Math.Min(70, celulaValor.Length)));
										break;
									case "CPF":
										dados[tituloColuna].Add(celulaValor.Contains(".") && celulaValor.Contains("-") && celulaValor.Length <= 14 ? celulaValor : celulaValor.Length == int.Parse(mascaraCPFLenth) ? Convert.ToUInt64(celulaValor).ToString(mascaraCPF) : "");
										break;
								}
							}
						}
					}
				}

				GravarExcel("asdf", dados);
			}

			catch (Exception error)
			{
				var mensagemErro = $"Falha na linha {indiceLinha}, coluna {colunaLetra}, Valor esperado: {tituloColuna}, valor da célula: \"{celulaValor}\": {error.Message}";

				if (!string.IsNullOrWhiteSpace(variaveisValor))
					mensagemErro += Environment.NewLine + "Variáveis" + Environment.NewLine + variaveisValor;

				throw new Exception(mensagemErro);
			}
		}

		private void GravarExcel(string nomeArquivo, Dictionary<string, List<object>> linhas)
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
				for (int i = 0; i < linha.Value.Count; i++)
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

	}
}
