using Migracao.Models;
using Migracao.Utils;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using System.Data;
using System.Text;

namespace Migracao.Sistems
{
    internal class OdontoCompany
	{
        string arquivoExcelCidades = "Files\\EnderecosCidades.xlsx";
		string arquivoExcelNomesUTF8 = "Files\\NomesUTF8.xlsx";

		string[] EMD101_Pacientes		= ["BAIRRO", "CEP", "CGC_CPF", "CIDADE", "CLIENTE", "CELULAR", "DT_CADASTRO", "DT_NASCIMENTO", "EMAIL", "ENDERECO", "ESTADO", "FONE1", "FONE2", "FORNECEDOR", "INSC_RG", "NOME", "NUM_CONVENIO", "NUM_ENDERECO", "NUM_FICHA", "OBS1", "SEXO_M_F"];
		string[] CRD111_Recebiveis		= ["AGENCIA", "AGUARDANDO_VINCULO", "ALINEA", "AXON_ID", "BANCO", "BANDA1", "BANDA2", "BANDA3", "BAIXA", "CAMPOX", "COD_CAIXA", "CODIGO_TUSS", "CONTA", "CONTA_CORRENTE", "CONTA_DOCUMENTO", "COBRADORA", "COBRANCA", "DATA_ENV_CART", "DATA_ENV_SCPC", "DATA_REMESSA", "DATA_RET_CART", "DATA_RET_SCPC", "DESCONTO_BOLETO", "DEVOLUCAO", "DOCUMENTO", "DT_AXON", "DUPLICATA", "EMISSAO", "EMITENTE", "ENCARGOS", "FILIAL", "GEROU_TRANSMISSAO", "GRUPO", "ID_BAIXAPLANOS", "ID_PIX", "JUROS", "LANCTO", "LOCAL", "LOJA", "MODIFICADO", "MOTIVO", "MULTA", "NOME_GRUPO", "NOME_LOCAL", "NOSSONUMERO", "NUM_BANCO", "OBS", "ORDEM", "PARCELA", "PERIODO", "PRAZO", "REAPRESENTOU", "RECEBEU_TRANSMISSAO", "REMESA", "RESPONSAVEL", "SEQ_ALINEA11", "SITUACAO", "SITUACAO_REMESSA", "TERMINAL", "TIPO_COBRANCA", "TIPO_DOC", "TOTAL", "TRANSMISSAO", "USUARIO", "VALOR", "VALOR_ORIG", "VALOR_RECEBER", "VALOR_VENDA", "VENCTO", "VENCTO_ORIG", "VR_CALCULADO", "VR_PARCELA"];
		string[] CXD555_Baixa			= ["AGENCIA", "BANCO", "BAIXA", "CALCULO", "CNPJ_CPF", "CONTA", "DATA", "DOCUMENTO", "DT_AXON", "DT_DEPOSITO", "FECHAR_DIRETO", "FICHA_FINANCEIRO", "HISTORICO", "HORA", "LANCTO", "LOJA", "LOTE", "MODIFICADO", "NUM_CONVENIO", "OBS1", "OBS2", "OBS3", "PERIODO", "PRO_MED", "PRO_ODO", "RESPONSAVEL", "ROY_MED", "ROY_ODO", "TERMINAL", "TIPO", "TRANSMISSAO", "USUARIO", "VALOR", "VALOR_RECEBER", "VLR_BRUTO"];
		string[] BXD111_Baixa			= ["AGUARDANDO_VINCULO", "AXON_ID", "BAIXA", "BANCO", "CAMPOX", "CGC_CPF", "COD_CAIXA", "CONTA_CORRENTE", "CONTA_DOCUMENTO", "DATA_REMESSA", "DOCUMENTO", "DT_AXON", "DUPLICATA", "GRUPO", "ID_BAIXAPLANOS", "LANCTO", "LOJA", "MODIFICADO", "MOTIVO", "NOME_GRUPO", "NUM_BANCO", "OBS", "PARCELA", "RESPONSAVEL", "TERMINAL", "TIPO_DOC", "TRANSMISSAO", "USUARIO", "VALOR", "VENCTO", "VR_CALCULADO", "VR_PARCELA"];
		string[] CED006_Dentistas		= ["ADMISSAO", "AGENCIA", "AGENCIA2", "AXON_ID", "BAIRRO", "BANCO", "BANCO2", "CAIXA_POSTAL", "CEP", "CGC_CPF", "CIDADE", "CLIENTE", "CODIGO", "CODIGO_CLIENTE", "CODIGO_INDICACAO", "CODIGO_VALIDADE", "COD_MUNICIPIO", "COD_PRAMELHOR", "COD_UF", "COD_VENDEDOR", "CONJUGE", "CONTA", "CONTA2", "CPF_FIA", "CPF_INDICACAO", "DATA_APROVACAO_DRCASH", "DATA_BLOQUEIO", "DATA_DEP_EXCLUIDO", "DATA_LGPD", "DATA_VALIDADE", "DEPENDENTE", "DT_AXON", "DT_CADASTRO", "DT_NASC_FIA", "DT_NASCIMENTO", "DT_NASCIMENTO_DEP", "DT_ULTMOV", "EMAIL", "ENDERECO", "ENDERECO_FIA", "ESTADO", "ESTADO_FIA", "FAX", "FONE1", "FONE2", "FONE_FIA", "FONE_REF_1", "FONE_REF_2", "F_OU_J", "FORNECEDOR", "FUNCAO", "ID_DRCASH", "INSC_RG", "INSTITUTO_ODC", "LGPD_CPF", "LGPD_DATA_HORA", "LGPD_IMAGEM", "LGPD_MENSAGEM", "LGPD_TELEFONE", "LGPD_USUARIO", "LOJA", "MAE", "MODIFICADO", "NOME", "NOME_FIA", "NOME_GRUPO", "NOME_LOCAL", "NOME_REF_1", "NOME_REF_2", "NOME_VALIDADE", "NUM_BLOQUEIO", "NUM_CONVENIO", "NUM_ENDERECO", "NUM_FICHA", "OBS1", "OBS_VALIDADE", "ONDE_TRABALHA", "PAI", "PARENTESCO_FIA", "PRESTADOR", "PROFISSAO", "PROFISSAO_FIA", "PROTETICO", "PROTETICO_ATIVO", "QTDE_DEPENDENTES", "RENDA_FIA", "RENDA_MES", "RG_FIA", "SEXO_M_F", "TITULAR", "TITULAR_DEP_EXCLUIDO", "TRANSMISSAO", "USU_BLOQUEIO", "USU_CADASTRO", "USUARIO", "USUARIO_LGPD", "USUARIO_VALIDADE", "VALOR_MAXIMO_DRCASH", "VR_LIMITE"];
		string[] CED001_Procedimentos	= ["CODIGO", "NOME", "SIMBOLO", "GRUPO", "CODIGO_TUSS", "VRVENDA", "ATIVO", "OBS", "PARTICULAR"];
		string[] CED002_CodProcedimentos = ["CODIGO", "NOME", "USUARIO"];

		List<string> cabecalhos_Pacientes		= ["Código", "Ativo(S/N)", "NomeCompleto", "NomeSocial", "Apelido", "Documento(CPF,CNPJ,CGC)", "DataCadastro(01/12/2024)", "Observações", "Email", "RG", "Sexo(M/F)", "NascimentoData", "NascimentoLocal", "EstadoCivil(S/C/V)", "Profissao", "CargoNaClinica", "Dentista(S/N)", "ConselhoCodigo", "Paciente(S/N)", "Funcionario(S/N)", "Fornecedor(S/N)", "TelefonePrincipal", "Celular", "TelefoneAlternativo", "Logradouro", "LogradouroNum", "Complemento", "Bairro", "Cidade", "Estado(SP)", "CEP(00000-000)"];
		List<string> cabecalhos_Recebiveis		= ["CPF", "Nome", "DocumentoRef", "RecebívelExigível(R/E)", "ValorOriginal", "ValorPago", "Prazo", "Vencimento(01/12/2010)", "DataBaixa", "Emissão(01/12/2010)", "ObservaçãoRecebível", "ObservaçãoRecebido"];
		List<string> cabecalhos_Agendamentos	= ["ID", "CPF", "Nome Completo", "Telefone", "Data Início (01/12/2024 00:00)", "Data Término (01/12/2024 00:00)", "Data Inclusão (01/12/2024)", "NomeCompletoDentista", "Observacao"];
		List<string> cabecalhos_Procedimentos	= ["Nome Tabela", "Ativo(S/N)", "Procedimento(Nome)", "Abreviação", "Especialidade", "Preço", "TUSS", "Diagnóstico(S/N)", "Prevenção(S/N)", "Odontopediatria(S/N)", "Dentística(S/N)", "Endodontia(S/N)", "Periodontia(S/N)", "Prótese(S/N)", "Cirurgia(S/N)", "Ortodontia(S/N)", "Radiologia(S/N)", "Estética(S/N)", "Implantodontia(S/N)", "Odontogeriatria(S/N)", "DTM(S/N)", "Orofacial(S/N)", ];
		List<string> cabecalhos_CodProcedimentos = ["ID", "Nome", "Usuário"];

		HashSet<string> cadastroPaciente, registroRecebivel;

		public Tuple<List<string[]>, List<string>> LerArquivosExcelCsv(string arquivo, Encoding encoding)
		{
			var separador = ExcelHelper.DetectarSeparadorCSV(arquivo);

			List<string> cabecalhosCSV = ExcelHelper.GetCabecalhosCSV(arquivo, separador, encoding);
			List<string[]> linhasCSV = ExcelHelper.GetLinhasCSV(arquivo, separador, cabecalhosCSV.Count(), encoding);

			return new Tuple<List<string[]>, List<string>>(linhasCSV, cabecalhosCSV);
		}

		public void LerArquivos(string estabelecimentoID, ListView listView = null)
		{
			ExcelHelper excelHelper = new();
			DataTable dataTablePessoas = new();
			DataTable dataTableRecebiveis = new();
			DataTable dataTableAgendamentos = new();
			DataTable dataTableProcedimentos = new();
			DataTable dataTableCodProcedimentos = new();
			//registroRecebivel = new HashSet<string>();

			foreach (string coluna in cabecalhos_Pacientes)
				dataTablePessoas.Columns.Add(coluna, typeof(string));

			foreach (string coluna in cabecalhos_Recebiveis)
				dataTableRecebiveis.Columns.Add(coluna, typeof(string));

			foreach (string coluna in cabecalhos_Agendamentos)
				dataTableAgendamentos.Columns.Add(coluna, typeof(string));

			if (File.Exists(arquivoExcelCidades))
				try
				{
					var workbookCidades = excelHelper.LerExcel(arquivoExcelCidades);
					var sheetCidades = workbookCidades.GetSheetAt(0);
					excelHelper.InitializeDictionaryCidade(sheetCidades);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelCidades}\": {ex.Message}");
				}			

			var excel_EMD101 = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("EMD101")));
			if (excel_EMD101 != null)
			{
				var resultado = LerArquivosExcelCsv(excel_EMD101.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;

				if (EMD101_Pacientes.All(cabecalhosCSV.Contains))
					dataTablePessoas = ConvertExcelPessoasPacientes(dataTablePessoas, cabecalhosCSV, linhasCSV);
			}

			var excel_CED006 = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CED006")));
			if (excel_CED006 != null)
			{
				var resultado = LerArquivosExcelCsv(excel_CED006.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;
				dataTablePessoas = ConvertExcelPessoasDentistas(dataTablePessoas, cabecalhosCSV, linhasCSV);

				if (CED006_Dentistas.All(cabecalhosCSV.Contains))
					dataTablePessoas = ConvertExcelPessoasPacientes(dataTablePessoas, cabecalhosCSV, linhasCSV);
			}

			if (excel_CED006 != null || excel_EMD101 != null)
			{
				var salvarArquivoPessoas = Tools.GerarNomeArquivo($"CadastroPessoas_{estabelecimentoID}_OdontoCompany");
				excelHelper.CriarExcelArquivo(salvarArquivoPessoas + ".xlsx", dataTablePessoas);
			}


			var excel_CRD111 = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CRD111")));
			if (excel_CRD111 != null)
			{
				var resultado = LerArquivosExcelCsv(excel_CRD111.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;
				dataTableRecebiveis = ConvertExcelRecebiveis(dataTableRecebiveis, cabecalhosCSV, linhasCSV, dataTablePessoas);
			}

			var excel_BXD111 = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("BXD111")));
			if (excel_BXD111 != null)
			{
				var resultado = LerArquivosExcelCsv(excel_BXD111.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;
				dataTableRecebiveis = ConvertExcelRecebidos(dataTableRecebiveis, cabecalhosCSV, linhasCSV, dataTablePessoas);
			}

			//var excel_CXD555 = listView.Items.Cast<ListViewItem>()
			//	.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CXD555")));
			//if (excel_CXD555 != null)
			//{
			//	var resultado = LerArquivosExcelCsv(excel_CXD555.Text, Encoding.UTF8);
			//	var linhasCSV = resultado.Item1;
			//	var cabecalhosCSV = resultado.Item2;
			//	dataTableRecebiveis = ConvertExcelRecebidos(dataTableRecebiveis, cabecalhosCSV, linhasCSV);
			//}

			if (excel_BXD111 != null || excel_CRD111 != null)
			{
				var salvarArquivoRecebiveis = Tools.GerarNomeArquivo($"CadastroRecebiveis_{estabelecimentoID}_OdontoCompany");
				excelHelper.CriarExcelArquivo(salvarArquivoRecebiveis + ".xlsx", dataTableRecebiveis);
			}


			var excel_AGENDA = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("AGENDA")));
			if (excel_AGENDA != null)
			{
				var resultado = LerArquivosExcelCsv(excel_AGENDA.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;
				dataTableAgendamentos = ConvertExcelAgendamento(dataTableAgendamentos, cabecalhosCSV, linhasCSV, dataTablePessoas);
			}

			if (excel_AGENDA != null)
			{
				var salvarArquivoAgenda = Tools.GerarNomeArquivo($"CadastroAgenda_{estabelecimentoID}_OdontoCompany");
				excelHelper.CriarExcelArquivo(salvarArquivoAgenda + ".xlsx", dataTableAgendamentos);
			}


			var excel_CED002 = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CED002")));
			if (excel_CED002 != null)
			{
				foreach (string coluna in cabecalhos_CodProcedimentos)
					dataTableCodProcedimentos.Columns.Add(coluna, typeof(string));

				var resultado = LerArquivosExcelCsv(excel_CED002.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;

				if (CED002_CodProcedimentos.All(cabecalhosCSV.Contains))
					dataTableCodProcedimentos = ConvertExcelGrupoProcedimentos(dataTableCodProcedimentos, cabecalhosCSV, linhasCSV);
			}

			var excel_CED001 = listView.Items.Cast<ListViewItem>()
				.FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CED001")));
			if (excel_CED001 != null)
			{
				foreach (string coluna in cabecalhos_Procedimentos)
					dataTableProcedimentos.Columns.Add(coluna, typeof(string));

				var resultado = LerArquivosExcelCsv(excel_CED001.Text, Encoding.UTF8);
				var linhasCSV = resultado.Item1;
				var cabecalhosCSV = resultado.Item2;

				if (CED001_Procedimentos.All(cabecalhosCSV.Contains))
					dataTableProcedimentos = ConvertExcelProcedimentos(dataTableProcedimentos, cabecalhosCSV, linhasCSV, dataTableCodProcedimentos);
			}

			if (dataTableProcedimentos.Rows.Count > 0)
			{
				var salvarArquivoAgenda = Tools.GerarNomeArquivo($"CadastroProcedimentos_{estabelecimentoID}_OdontoCompany");
				excelHelper.CriarExcelArquivo(salvarArquivoAgenda + ".xlsx", dataTableProcedimentos);
			}
		}

		public DataTable ConvertExcelRecebiveis(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePacientes)
		{
			try
			{
				int linhaIndex = 0;
				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length)
								valoresLinha.Add(cabecalhos[i], linha[i]);

						var cpf = valoresLinha.GetValueOrDefault("CGC_CPF").Trim();
						var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
						var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();
						var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
						var vencimentoData = valoresLinha.GetValueOrDefault("VENCTO").Trim();
						var emissaoData = valoresLinha.GetValueOrDefault("EMISSAO").Trim();

						cpf = cpf.ToCPF();
						string nome = "";

						if (dataTablePacientes.Rows.Count > 0)
						{
							DataRow[] dataRowEncontrados = dataTablePacientes.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
							if (dataRowEncontrados.Length > 0)
								nome = dataRowEncontrados[0]["NomeCompleto"].ToString();
						}

						if (valor.ArredondarValorV2() > 1)
						{
							dataRow["CPF"] = cpf;
							dataRow["Nome"] = nome;
							dataRow["ObservaçãoRecebível"] = observacao;
							dataRow["DocumentoRef"] = documento;
							dataRow["ValorOriginal"] = valor.ArredondarValorV2();
							//dataRow["Prazo"] = prazo;
							dataRow["Vencimento(01/12/2010)"] = vencimentoData.ToData().ToString("dd/MM/yyyy");
							dataRow["Emissão(01/12/2010)"] = emissaoData.ToData().ToString("dd/MM/yyyy");
							dataRow["RecebívelExigível(R/E)"] = "R";

							dataTable.Rows.Add(dataRow);
						}
					}
					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
			}
		}

		public DataTable ConvertExcelRecebidos(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePacientes)
		{
			try
			{
				int linhaIndex = 0;
				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length)
								valoresLinha.Add(cabecalhos[i], linha[i]);

						var cabecalhoCpf = "CGC_CPF";
						var cabecalhoObs = "OBS";
						var cabecalhoDataBaixa = "BAIXA";
						var cabecalhoDataVencimento = "VENCTO";

						if (cabecalhos.Contains("CNPJ_CPF"))
						{
							cabecalhoCpf = "CNPJ_CPF";
							cabecalhoObs = "OBS1";
							cabecalhoDataBaixa = "DATA";
							cabecalhoDataVencimento = "DATA";
						}

						var cpf = valoresLinha.GetValueOrDefault(cabecalhoCpf).Trim();
						var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();
						var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
						var valorOriginal = valoresLinha.GetValueOrDefault("VR_PARCELA").Trim();
						var observacao = valoresLinha.GetValueOrDefault(cabecalhoObs).Trim();
						var baixaData = valoresLinha.GetValueOrDefault(cabecalhoDataBaixa).Trim();
						var vencimentoData = valoresLinha.GetValueOrDefault(cabecalhoDataVencimento).Trim();


						cpf = cpf.ToCPF();
						string nome = "";

						if (dataTablePacientes.Rows.Count > 0)
						{
							DataRow[] dataRowPessoasEncontrados = dataTablePacientes.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
							if (dataRowPessoasEncontrados.Length > 0)
								nome = dataRowPessoasEncontrados[0]["NomeCompleto"].ToString();
						}


						DataRow[] dataRowEncontrados = dataTable.AsEnumerable()
						.Where(row =>
							row.Field<string>("DocumentoRef") == documento &&
							row.Field<string>("CPF") == cpf &&
							row.Field<string>("ValorOriginal") == valorOriginal.ArredondarValorV2().ToString())
						.ToArray();

						if (dataRowEncontrados.Length > 0)
						{
							foreach (var dataRowEncontrado in dataRowEncontrados)
							{
								dataRowEncontrado["ValorPago"] = valor.ArredondarValorV2();
								dataRowEncontrado["DataBaixa"] = baixaData.ToData().ToString("dd/MM/yyyy");
								dataRowEncontrado["ObservaçãoRecebido"] = observacao;
							}
						}

						else
						{
							dataRow["CPF"] = cpf;
							dataRow["Nome"] = nome;
							dataRow["DocumentoRef"] = documento;
							dataRow["ValorPago"] = valor.ArredondarValorV2();
							dataRow["DataBaixa"] = baixaData;
							dataRow["Vencimento(01/12/2010)"] = vencimentoData.ToData();
							dataRow["ObservaçãoRecebido"] = observacao;
							dataRow["RecebívelExigível(R/E)"] = "R";

							dataTable.Rows.Add(dataRow);
						}
					}
					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel Recebidos: {error.Message}");
			}
		}

		public DataTable ConvertExcelPessoasDentistas(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas)
		{
			try
			{
				int linhaIndex = 0;
				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length) // Verificar se o índice está dentro do tamanho da linha
								valoresLinha.Add(cabecalhos[i], linha[i]);

						//if (fichasCadastradas.Contains(codigo.ToNum()))
						var codigo = valoresLinha.GetValueOrDefault("CODIGO").Trim();
						var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
						var departamento = valoresLinha.GetValueOrDefault("DEPARTAMENTO").Trim();
						var obs = valoresLinha.GetValueOrDefault("OBS").Trim();
						var ativo = valoresLinha.GetValueOrDefault("ATIVO").Trim();
						var nomeCompleto = valoresLinha.GetValueOrDefault("NOME_COMPLETO").Trim();
						var email = valoresLinha.GetValueOrDefault("EMAIL").Trim();
						var telefone = valoresLinha.GetValueOrDefault("TELEFONE").Trim();
						var cro = valoresLinha.GetValueOrDefault("CRO").Trim();
						var modificado = valoresLinha.GetValueOrDefault("MODIFICADO").Trim();

						dataRow["Código"] = codigo.ToNum();
						dataRow["Ativo(S/N)"] = "S";
						dataRow["NomeCompleto"] = nome.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
						dataRow["NomeSocial"] = "";
						dataRow["Apelido"] = nome.GetLetras().GetPrimeirosCaracteres(19).PrimeiraLetraMaiuscula();
						//dataRow["Documento(CPF,CNPJ,CGC)"] = cgcCpf.ToCPF();
						dataRow["DataCadastro(01/12/2024)"] = modificado.ToData();
						dataRow["Observações"] = obs;
						dataRow["Email"] = email.ToEmail();
						//dataRow["RG"] = rg.GetPrimeirosCaracteres(20);
						//dataRow["Sexo(M/F)"] = sexo.ToSexo("m", "f").ToSN();
						//dataRow["NascimentoData"] = dataNascimento.ToData();
						dataRow["NascimentoLocal"] = "";
						dataRow["EstadoCivil(S/C/V)"] = "";
						dataRow["Profissao"] = "";
						dataRow["CargoNaClinica"] = "";
						dataRow["Dentista(S/N)"] = "N";
						dataRow["ConselhoCodigo"] = "";
						dataRow["Paciente(S/N)"] = "N";
						dataRow["Funcionario(S/N)"] = "S";
						dataRow["Fornecedor(S/N)"] = "N";

						dataTable.Rows.Add(dataRow);
					}
					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel para Pessoas Dentistas: {error.Message}");
			}
		}

		public DataTable ConvertExcelPessoasPacientes(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas)
		{
			ExcelHelper excelHelper = new();
			try
			{
				int linhaIndex = 0;
				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length) // Verificar se o índice está dentro do tamanho da linha
								valoresLinha.Add(cabecalhos[i], linha[i]);

						var numFicha = valoresLinha.GetValueOrDefault("NUM_FICHA").Trim();
						var cliente = valoresLinha.GetValueOrDefault("CLIENTE").Trim();
						var fornecedor = valoresLinha.GetValueOrDefault("FORNECEDOR").Trim();
						var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
						var cgcCpf = valoresLinha.GetValueOrDefault("CGC_CPF").Trim();
						var rg = valoresLinha.GetValueOrDefault("INSC_RG").Trim();
						var sexo = valoresLinha.GetValueOrDefault("SEXO_M_F").Trim();
						var email = valoresLinha.GetValueOrDefault("EMAIL").Trim();
						var fone1 = valoresLinha.GetValueOrDefault("FONE1").Trim();
						var fone2 = valoresLinha.GetValueOrDefault("FONE2").Trim();
						var celular = valoresLinha.GetValueOrDefault("CELULAR").Trim();
						var endereco = valoresLinha.GetValueOrDefault("ENDERECO").Trim();
						var bairro = valoresLinha.GetValueOrDefault("BAIRRO").Trim();
						var numEndereco = valoresLinha.GetValueOrDefault("NUM_ENDERECO").Trim();
						var cidade = valoresLinha.GetValueOrDefault("CIDADE").Trim();
						var estado = valoresLinha.GetValueOrDefault("ESTADO").Trim();
						var cep = valoresLinha.GetValueOrDefault("CEP").Trim();
						var obs = valoresLinha.GetValueOrDefault("OBS1").Trim();
						var numConvenio = valoresLinha.GetValueOrDefault("NUM_CONVENIO").Trim();
						var dataCadastro = valoresLinha.GetValueOrDefault("DT_CADASTRO").Trim();
						var dataNascimento = valoresLinha.GetValueOrDefault("DT_NASCIMENTO").Trim();

						if (cliente != "S" && fornecedor != "S")
							cliente = "S";

						if (!excelHelper.CidadeExists(cidade.PrimeiraLetraMaiuscula(), estado))
							cidade = cidade.EncontrarCidadeSemelhante();

						//dataRow["Codigo"] = numFicha.ToNum();
						dataRow["Ativo(S/N)"] = "S";
						dataRow["NomeCompleto"] = nome.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
						dataRow["NomeSocial"] = "";
						dataRow["Apelido"] = nome.GetLetras().GetPrimeirosCaracteres(19).PrimeiraLetraMaiuscula();
						dataRow["Documento(CPF,CNPJ,CGC)"] = cgcCpf.ToCPF();
						dataRow["DataCadastro(01/12/2024)"] = dataCadastro.ToData();
						dataRow["Observações"] = obs;
						dataRow["Email"] = email.ToEmail();
						dataRow["RG"] = rg.GetPrimeirosCaracteres(20);
						dataRow["Sexo(M/F)"] = sexo.ToSexo("m", "f").ToSN();
						dataRow["NascimentoData"] = dataNascimento.ToData();
						dataRow["NascimentoLocal"] = "";
						dataRow["EstadoCivil(S/C/V)"] = "";
						dataRow["Profissao"] = "";
						dataRow["CargoNaClinica"] = "";
						dataRow["Dentista(S/N)"] = "N";
						dataRow["ConselhoCodigo"] = "";
						dataRow["Paciente(S/N)"] = cliente;
						dataRow["Funcionario(S/N)"] = "N";
						dataRow["Fornecedor(S/N)"] = fornecedor;
						dataRow["TelefonePrincipal"] = fone1.ToFone();
						dataRow["Celular"] = celular.ToFone();
						dataRow["TelefoneAlternativo"] = fone2.ToFone();
						dataRow["Logradouro"] = endereco.PrimeiraLetraMaiuscula();
						dataRow["LogradouroNum"] = numEndereco;
						dataRow["Complemento"] = "";
						dataRow["Bairro"] = bairro.PrimeiraLetraMaiuscula();
						dataRow["Cidade"] = cidade.EncontrarCidadeSemelhante();
						dataRow["Estado(SP)"] = estado.ToUpper();
						dataRow["CEP(00000-000)"] = cep.ToNum();

						dataTable.Rows.Add(dataRow);
					}

					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
			}
		}

		public DataTable ConvertExcelGrupoProcedimentos(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas)
		{
			try
			{
				int linhaIndex = 0;
				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length)
								valoresLinha.Add(cabecalhos[i], linha[i]);

						var cod = valoresLinha.GetValueOrDefault("CODIGO").Trim();
						var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
						var usuario = valoresLinha.GetValueOrDefault("USUARIO").Trim();

						dataRow["ID"] = cod;
						dataRow["Nome"] = nome.GetPrimeirosCaracteres(100).PrimeiraLetraMaiuscula();
						dataRow["Usuário"] = nome.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();

						dataTable.Rows.Add(dataRow);
					}

					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel Grupo Procedimentos: {error.Message}");
			}
		}

		public DataTable ConvertExcelProcedimentos(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable codProcedimentos)
		{
			try
			{
				var grupoCategoriaDict = new Dictionary<string, string>()
				{
					{ "CIRURGIA", "Cirurgia" },
					{ "ENDODONTIA", "Endodontia" },
					{ "PERIODONTIA", "Periodontia" },
					{ "PROTESE", "Protese" },
					{ "CLINICO", "Outros" },
					{ "MANUTENCAO", "Ortodontia" },
					{ "ORTODONTIA", "Ortodontia" },
					{ "ORTO", "Ortodontia" },
					{ "PREVENCAO", "Prevencao" },
					{ "OROFACIAL", "Orofacial" },
					{ "HARMONIZACAO OROFACIAL", "Orofacial" }
				};
				int linhaIndex = 0;

				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length)
								valoresLinha.Add(cabecalhos[i], linha[i]);

						var cod = valoresLinha.GetValueOrDefault("GRUPO").Trim();
						var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
						var abreviacao = valoresLinha.GetValueOrDefault("SIMBOLO").Trim();
						var tuss = valoresLinha.GetValueOrDefault("CODIGO_TUSS").Trim();
						var valor = valoresLinha.GetValueOrDefault("VRVENDA").Trim();
						var ativo = valoresLinha.GetValueOrDefault("ATIVO").Trim();
						var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
						var particular = valoresLinha.GetValueOrDefault("PARTICULAR").Trim();
						var nomeTabela = "";
						var especialidade = "";

						DataRow[] dataRowEncontrados = codProcedimentos.AsEnumerable()
						.Where(row =>
							row.Field<string>("ID") == cod)
						.ToArray();

						//DataRow[] dataRowEncontrados = codProcedimentos.Select($"ID = '{cod}'");
						if (dataRowEncontrados.Length > 0)
							especialidade = dataRowEncontrados[0]["Nome"].ToString();

						if (nome.StartsWith("odc ", StringComparison.CurrentCultureIgnoreCase))
							nomeTabela = "ODC";
						//else if (string.IsNullOrEmpty(observacao) && particular != "N")
						else if (string.IsNullOrEmpty(observacao) || grupoCategoriaDict.ContainsKey(Tools.RemoverAcentos(especialidade).ToUpper()))
							nomeTabela = "Particular";
						else
						{
							nomeTabela = especialidade;
							especialidade = "Outros";
						}

						var buscarEspecialidade = Tools.RemoverAcentos(especialidade).ToUpper();

						if (grupoCategoriaDict.ContainsKey(buscarEspecialidade))
							especialidade = grupoCategoriaDict[buscarEspecialidade];

						if (nomeTabela.GetPrimeirosCaracteres(40).PrimeiraLetraMaiuscula() == "")
							nomeTabela = nomeTabela;

						dataRow["Nome Tabela"] = nomeTabela.GetPrimeirosCaracteres(40).PrimeiraLetraMaiuscula();
						dataRow["Especialidade"] = especialidade;
						dataRow["Ativo(S/N)"] = ativo == "N" ? "N" : "S";
						dataRow["Procedimento(Nome)"] = nome.GetPrimeirosCaracteres(100).PrimeiraLetraMaiuscula();
						dataRow["Abreviação"] = abreviacao;
						dataRow["Preço"] = valor.ArredondarValorV2();
						dataRow["TUSS"] = tuss.ToNumV2();

						dataTable.Rows.Add(dataRow);
					}

					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel Procedimentos: {error.Message}");
			}
		}

		public DataTable ConvertExcelAgendamento(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
		{
			List<string> ids = new();

			try
			{
				int linhaIndex = 0;
				foreach (string[] linha in linhas)
				{
					try
					{
						DataRow dataRow = dataTable.NewRow();
						var valoresLinha = new Dictionary<string, string>();

						for (int i = 0; i < cabecalhos.Count; i++)
							if (i < linha.Length) // Verificar se o índice está dentro do tamanho da linha
								valoresLinha.Add(cabecalhos[i], linha[i]);

						var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
						var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
						var data = valoresLinha.GetValueOrDefault("DATA").Trim();
						var hora = valoresLinha.GetValueOrDefault("HORA").Trim();
						var cod_responsavel = valoresLinha.GetValueOrDefault("CODIGO_RESP").Trim();
						var responsavel = valoresLinha.GetValueOrDefault("RESPONSAVEL").Trim();
						var telefone = valoresLinha.GetValueOrDefault("FONE_1").Trim();
						var dataInclusao = valoresLinha.GetValueOrDefault("MODIFICADO").Trim();
						var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
						var id = valoresLinha.GetValueOrDefault("LANCTO").Trim();

						if (!ids.Contains(id))
						{
							ids.Add(id);

							DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable()
								.Where(row =>
								row.Field<string>("Código") == cod_responsavel)
								.ToArray();

							//DataRow[] dataRowEncontrados = dataTablePessoas.Select($"Código = '{cod_responsavel}'");
							
							if (dataRowEncontrados.Length > 0)
								responsavel = dataRowEncontrados[0]["NomeCompleto"].ToString();

							var minutos = hora.Split(':')[1];
							var horas = hora.Split(':')[0];
							var dataInicio = data.ToData();

							if (!string.IsNullOrEmpty(horas))
								dataInicio = dataInicio.AddHours(double.Parse(horas));
							if (!string.IsNullOrEmpty(minutos))
								dataInicio = dataInicio.AddMinutes(double.Parse(minutos));

							var dataTermino = dataInicio;
							var idsEncontrados = linhas.Where(linha => linha[0].Equals(id)).Count();
							if (idsEncontrados > 0)
								dataTermino = dataTermino.AddMinutes(15 * idsEncontrados);
							else
								dataTermino = dataTermino.AddMinutes(15);

							dataRow["ID"] = id;
							dataRow["CPF"] = cpf.ToCPF();
							dataRow["Nome Completo"] = nome.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
							dataRow["Data Início (01/12/2024 00:00)"] = dataInicio;
							dataRow["Data Término (01/12/2024 00:00)"] = dataTermino;
							dataRow["Data Inclusão (01/12/2024)"] = dataInclusao.ToData();
							dataRow["NomeCompletoDentista"] = responsavel;
							dataRow["Telefone"] = telefone.ToFone();
							dataRow["Observacao"] = observacao;

							dataTable.Rows.Add(dataRow);
						}
					}

					catch (Exception error)
					{
						throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
					}

					linhaIndex++;
				}

				return dataTable;
			}
			catch (Exception error)
			{
				throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
			}
		}

		public void ImportarAgenda(string arquivoExcel, int estabelecimentoID, string arquivoExcelAgenda, int loginID)
		{
			var dataHoje = DateTime.Now;
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoExcelAgenda))
			{
				ISheet sheet;
				try
				{
					IWorkbook workbook = excelHelper.LerExcel(arquivoExcelAgenda);
					sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionaryAgendamentos(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelAgenda}\": {ex.Message}");
				}
			}

			var agendamentos = new List<Agendamento>();

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", cpf = "", dentistaResponsavel = "";
					bool faltou = false;
					string? titulo = null, observacoes = null, documento = null;
					int recibo = 0, codigo = 0;
					int? consumidorID = null, fornecedorID = null, colaboradorID = null, funcionarioID = null, clienteID = null, pessoaID = null;
					decimal pagoValor = 0, valor = 0;
					long? telefone = null;
					byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
					DateTime dataConsulta = dataHoje, dataInclusao = dataHoje;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "CPF":
										cpf = celulaValor.ToCPF();
										break;
									case "NomeCompleto":
										nomeCompleto = celulaValor;
										break;
									case "Telefone":
										telefone = celulaValor.ToFone();
										break;
									case "DataHoraConsulta(01/12/2024 00:00)":
										dataConsulta = celulaValor.ToData();
										break;
									case "NomeCompletoDentista":
										dentistaResponsavel = celulaValor;
										break;
									case "Observacao":
										observacoes = celulaValor;
										break;
									case "DataInclusao(01/12/2024)":
										dataInclusao = celulaValor.ToData();
										break;
								}
							}
						}
					}

					if (dataConsulta == dataHoje)
						dataConsulta = dataInclusao;

					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: cpf, codigo: codigo.ToString());
					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: dentistaResponsavel);
					var funcionarioIDValue = excelHelper.GetFuncionarioID(nomeCompleto: dentistaResponsavel);

					if (!string.IsNullOrEmpty(consumidorIDValue))
						consumidorID = int.Parse(consumidorIDValue);
					if (!string.IsNullOrEmpty(pessoaIDValue))
						pessoaID = int.Parse(pessoaIDValue);
					if (!string.IsNullOrEmpty(funcionarioIDValue))
						funcionarioID = int.Parse(funcionarioIDValue);
					else
						titulo = nomeCompleto;

					if (!excelHelper.AgendamentoExists(nomeCompleto, dataConsulta, consumidorID))
					{
						if (telefone == null)
							telefone = 0;

						if (!string.IsNullOrEmpty(funcionarioIDValue) && !string.IsNullOrEmpty(consumidorIDValue) && !string.IsNullOrEmpty(pessoaIDValue))
							agendamentos.Add(new Agendamento()
							{
								LoginID = loginID,
								EstabelecimentoID = estabelecimentoID,
								AtendeTipoID = 1,
								DataInicio = dataConsulta,
								DataTermino = dataConsulta.AddMinutes(15),
								ConsumidorID = (int)consumidorID,
								ConsumidorPessoaNome = nomeCompleto,
								Titulo = titulo,
								FuncionarioID = (int)funcionarioID,
								DataInclusao = dataInclusao,
								ConsumidorPessoaFone1 = (long)telefone,
								Descricao = observacoes,
								PessoaID = (int)pessoaID
								//ConsumidorPessoaID = (int)pessoaID,
								//DataCancelamento = ,
								//
								//AtendimentoValor = ,
								//SecretariaID = ,
								//SalaID = ,
							});
					}

					else
						telefone = 0;
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "LoginID", agendamentos.ConvertAll(agendamento => (object)agendamento.LoginID).ToArray() },
					{ "EstabelecimentoID", agendamentos.ConvertAll(agendamento => (object)agendamento.EstabelecimentoID).ToArray() },
					{ "AtendeTipoID", agendamentos.ConvertAll(agendamento => (object)agendamento.AtendeTipoID).ToArray() },
					{ "DataInicio", agendamentos.ConvertAll(agendamento => (object)agendamento.DataInicio).ToArray() },
					{ "DataTermino", agendamentos.ConvertAll(agendamento => (object)agendamento.DataTermino).ToArray() },
					{ "ConsumidorID", agendamentos.ConvertAll(agendamento => (object)agendamento.ConsumidorID).ToArray() },
					{ "ConsumidorPessoaNome", agendamentos.ConvertAll(agendamento => (object)agendamento.ConsumidorPessoaNome).ToArray() },
					{ "Titulo", agendamentos.ConvertAll(agendamento => (object)agendamento.Titulo).ToArray() },
					{ "FuncionarioID", agendamentos.ConvertAll(agendamento => (object)agendamento.FuncionarioID).ToArray() },
					{ "DataInclusao", agendamentos.ConvertAll(agendamento => (object)agendamento.DataInclusao).ToArray() },
					{ "ConsumidorPessoaFone1", agendamentos.ConvertAll(agendamento => (object)agendamento.ConsumidorPessoaFone1).ToArray() },
					{ "Descricao", agendamentos.ConvertAll(agendamento => (object)agendamento.Descricao).ToArray() },
					{ "PessoaID", agendamentos.ConvertAll(agendamento => (object)agendamento.PessoaID).ToArray() }
				};

				var salvarArquivo = Tools.GerarNomeArquivo($"Agendamentos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Agendamentos", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);

				MessageBox.Show("Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarPrecosTabelas(string arquivoExcel, int estabelecimentoID, int loginID, string arquivoTabelaPrecosAtuais)
		{
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();
			var adicionados = new List<string>();
			var precosTabelas = new List<PrecosTabela>();
			ISheet sheet = null;

			if (!string.IsNullOrEmpty(arquivoTabelaPrecosAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoTabelaPrecosAtuais);
					sheet = workbook.GetSheetAt(0);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoTabelaPrecosAtuais}\": {ex.Message}");
				}


			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string? nome = null, abreviacao = null;
					decimal valor = 0;
					string especialidade = "Outros";
					long tuss = 0;
					bool ativo = true;
					byte categoria = (byte)ProcedimentosCategoriasID.Outros;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "Nome Tabela":
										nome = celulaValor.PrimeiraLetraMaiuscula();
										break;
								}
							}
						}
					}

					if (!adicionados.Contains(nome) && (!excelHelper.ExisteTexto(sheet, "Nome", nome) && !excelHelper.ExisteTexto(sheet, "Nome", "Migração - " + nome)))
					{
						adicionados.Add(nome);

						precosTabelas.Add(new PrecosTabela
						{
							Ativo = ativo,
							DataInclusao = DateTime.Now,
							LoginID = loginID,
							SeguimentoID = 1,
							SolucaoID = 1,
							EstabelecimentoID = estabelecimentoID,
							Nome = nome
						});
					}
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "Ativo", precosTabelas.ConvertAll(precotabela => (object)precotabela.Ativo).ToArray() },
					{ "DataInclusao", precosTabelas.ConvertAll(precotabela => (object)precotabela.DataInclusao).ToArray() },
					{ "LoginID", precosTabelas.ConvertAll(precotabela => (object)precotabela.LoginID).ToArray() },
					{ "SeguimentoID", precosTabelas.ConvertAll(precotabela => (object)precotabela.SeguimentoID).ToArray() },
					{ "SolucaoID", precosTabelas.ConvertAll(precotabela => (object)precotabela.SolucaoID).ToArray() },
					{ "EstabelecimentoID", precosTabelas.ConvertAll(precotabela => (object)precotabela.EstabelecimentoID).ToArray() },
					{ "Nome", precosTabelas.ConvertAll(precotabela => (object)precotabela.Nome).ToArray() }
				};

				var salvarArquivo = Tools.GerarNomeArquivo($"PrecosTabelas_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("PrecosTabelas", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);

				MessageBox.Show("Limpar o Redis" + Environment.NewLine + "redis-cli.exe -h 127.0.0.1 -n 0 del Tabelas:Tabelas-" + estabelecimentoID.ToString("D6"), "Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarPrecos(string arquivoExcel, int estabelecimentoID, int loginID, string arquivoTabelaPrecosAtuais)
        {
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			var grupoCategoriaDict = new Dictionary<string, ProcedimentosCategoriasID>()
			{
				{ "CIRURGIA", ProcedimentosCategoriasID.Cirurgia },
				{ "ENDODONTIA", ProcedimentosCategoriasID.Endodontia },
				{ "PERIODONTIA", ProcedimentosCategoriasID.Periodontia },
				{ "PROTESE", ProcedimentosCategoriasID.Prótese },
				{ "CLINICO", ProcedimentosCategoriasID.Outros },
				{ "MANUTENCAO", ProcedimentosCategoriasID.Ortodontia },
				{ "ORTODONTIA", ProcedimentosCategoriasID.Ortodontia },
				{ "PREVENCAO", ProcedimentosCategoriasID.Prevenção },
				{ "OROFACIAL", ProcedimentosCategoriasID.Orofacial },
				{ "OUTROS", ProcedimentosCategoriasID.Outros }
			};

			var precosTabelas = new List<PrecosTabela>();
			var precos = new List<Preco>();
			ISheet sheet = null;

			if (!string.IsNullOrEmpty(arquivoTabelaPrecosAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoTabelaPrecosAtuais);
					sheet = workbook.GetSheetAt(0);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoTabelaPrecosAtuais}\": {ex.Message}");
				}

			try
            {
                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

					string? titulo = null, abreviacao = null;
                    decimal valor = 0;
                    string especialidade = "Outros", nomeTabela = "";
					long tuss = 0;
					bool ativo = true;
                    byte categoria = (byte)ProcedimentosCategoriasID.Outros;

                    foreach (var celula in linha.Cells)
                    {
                        if (celula != null)
                        {
                            celulaValor = celula.ToString().Trim().Replace("'", "’");
                            tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
                            colunaLetra = excelHelper.GetColumnLetter(celula);

                            if (!string.IsNullOrWhiteSpace(celulaValor))
                            {
                                switch (tituloColuna)
                                {
                                    case "Nome Tabela":
                                        nomeTabela = celulaValor.PrimeiraLetraMaiuscula();
                                        break;
                                    case "Ativo(S/N)":
										ativo = celulaValor == "S" ? true : false;
                                        break;
                                    case "Procedimento(Nome)":
										titulo = celulaValor;
                                        break;
									case "Abreviação":
										abreviacao = celulaValor;
										break;
									case "Especialidade":
										especialidade = celulaValor;
										break;
									case "Preço":
										valor = celulaValor.ToMoeda();
										break;
									case "TUSS":
										tuss = celulaValor.ToNumV2();
										break;
								}
                            }
                        }
                    }


					if (!excelHelper.ExisteTexto(sheet, "Titulo", titulo)) {

						var tabelaIdEcontrada = excelHelper.ProcurarCelula(sheet, "Nome", nomeTabela, "TabelaID");
						if (string.IsNullOrEmpty(tabelaIdEcontrada))
							tabelaIdEcontrada = excelHelper.ProcurarCelula(sheet, "Nome", "Migração - " + nomeTabela, "TabelaID");

						if (string.IsNullOrEmpty(tabelaIdEcontrada))
						{
							if (grupoCategoriaDict.ContainsKey(especialidade))
								categoria = (byte)grupoCategoriaDict[especialidade];

							precos.Add(new Preco
							{
								Ativo = ativo,
								CategoriaID = categoria,
								DataInclusao = DateTime.Now,
								LoginID = loginID,
								TabelaID = int.Parse(tabelaIdEcontrada),
								Titulo = titulo,
								Valor = valor,
								CodigoTISS = tuss,
								Atalho = abreviacao
							});
						}
					}
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "LoginID", precos.ConvertAll(preco => (object)preco.LoginID).ToArray() },
					{ "Ativo", precos.ConvertAll(preco => (object)preco.Ativo).ToArray() },
					{ "CategoriaID", precos.ConvertAll(preco => (object)preco.CategoriaID).ToArray() },
					{ "DataInclusao", precos.ConvertAll(preco => (object)preco.DataInclusao).ToArray() },
					{ "TabelaID", precos.ConvertAll(preco => (object)preco.TabelaID).ToArray() },
					{ "Titulo", precos.ConvertAll(preco => (object)preco.Titulo).ToArray() },
					{ "Valor", precos.ConvertAll(preco => (object)preco.Valor).ToArray() },
					{ "CodigoTISS", precos.ConvertAll(preco => (object)preco.CodigoTISS).ToArray() },
					{ "Atalho", precos.ConvertAll(preco => (object)preco.Atalho).ToArray() }
				};

				var salvarArquivo = Tools.GerarNomeArquivo($"Precos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Precos", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);

				MessageBox.Show("Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarRecebiveis(string arquivoExcel, string arquivoExcelRecebiveisProd, int estabelecimentoID, int respFinanceiroPessoaID, int loginID)
        {
			var dataHoje = DateTime.Now;
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			ISheet sheetConsumidores;
			try
			{
				IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelRecebiveisProd);
				sheetConsumidores = workbookConsumidores.GetSheetAt(0);
				excelHelper.InitializeDictionaryRecebiveis(sheetConsumidores);
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelRecebiveisProd}\": {ex.Message}");
			}

            //var excelRecebidosDict = ExcelRecebidosToDictionary(arquivoExcelBaixa);

			var fluxoCaixas = new List<FluxoCaixa>();
			var recebiveis = new List<Recebivel>();

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", cpf = "";
                    string? outroSacadoNome = null, observacoes = null, documento = null;
					int recibo = 0, codigo = 0;
					int? consumidorID = null, fornecedorID = null, colaboradorID = null, funcionarioID = null, clienteID = null;
					decimal pagoValor = 0, valor = 0;
					bool recebivel = false, baixa = false;
					byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
					DateTime dataPagamento = dataHoje, nascimentoData = dataHoje, dataVencimento = dataHoje, dataInclusao = dataHoje;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "CPF":
										cpf = celulaValor.ToCPF();
										break;
									case "Emitente":
										outroSacadoNome = celulaValor;
										break;
									case "DocumentoRef":
										documento = celulaValor;
										break;
									case "Vencimento(01/12/2010)":
										dataVencimento = celulaValor.ToData();
										break;
									case "Emissão(01/12/2010)":
										dataInclusao = celulaValor.ToData();
										break;
									case "ObservaçãoRecebível":
                                        observacoes = celulaValor;//.ToTipoPagamento()
										break;
									case "ValorOriginal":
										valor = celulaValor.ArredondarValor();
										break;
									case "RecebívelExigível(R/E)":
										recebivel = celulaValor == "R" ? true : false;
										break;
									case "DataBaixa":
										baixa = !string.IsNullOrEmpty(celulaValor) ? true : false;
										break;
								}
							}
						}
					}

					if (recebivel && !baixa)
					{
						if (dataVencimento == dataHoje)
							dataVencimento = new DateTime(dataInclusao.Year, dataVencimento.Month, dataVencimento.Day);

						var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: cpf, codigo: codigo.ToString());
						var fornecedorIDValue = excelHelper.GetFornecedorID(nomeCompleto: nomeCompleto, cpf: cpf);
						var funcionarioIDValue = excelHelper.GetFuncionarioID(nomeCompleto: nomeCompleto, cpf: cpf);

						if (!string.IsNullOrEmpty(consumidorIDValue))
							consumidorID = int.Parse(consumidorIDValue);
						else if (!string.IsNullOrEmpty(fornecedorIDValue))
							fornecedorID = int.Parse(fornecedorIDValue);
						else if (!string.IsNullOrEmpty(funcionarioIDValue))
							funcionarioID = int.Parse(funcionarioIDValue);
						else
							outroSacadoNome = cpf;

						if (consumidorID == 18283648)
							consumidorID = consumidorID;

						if (!string.IsNullOrEmpty(consumidorIDValue))
						{
							if (!excelHelper.RecebivelExists((int)consumidorID, valor, dataVencimento))
								recebiveis.Add(new Recebivel()
								{
									ConsumidorID = consumidorID,
									FornecedorID = fornecedorID,
									ClienteID = clienteID,
									ColaboradorID = colaboradorID,
									SacadoNome = outroSacadoNome,
									EspecieID = (byte)formaPagamento,
									DataEmissao = dataInclusao,
									ValorOriginal = valor,
									ValorDevido = valor,
									DataBaseCalculo = dataInclusao,
									DataInclusao = dataInclusao,
									DataVencimento = dataVencimento,
									FinanceiroID = respFinanceiroPessoaID,
									LoginID = loginID,
									EstabelecimentoID = estabelecimentoID,
									SituacaoID = (byte)TituloSituacoesID.Normal,
									Observacoes = observacoes,
									ExclusaoMotivo = documento
									//OrcamentoID
									//PlanoContasID
									//Documento = contratoControle
									//ContratoID = contratoID
								});
						}
					}
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "ConsumidorID", recebiveis.ConvertAll(recebivel => (object)recebivel.ConsumidorID).ToArray() },
                    { "FornecedorID", recebiveis.ConvertAll(recebivel => (object)recebivel.FornecedorID).ToArray() },
                    { "ClienteID", recebiveis.ConvertAll(recebivel => (object)recebivel.ClienteID).ToArray() },
                    { "ColaboradorID", recebiveis.ConvertAll(recebivel => (object)recebivel.ColaboradorID).ToArray() },
                    { "SacadoNome", recebiveis.ConvertAll(recebivel => (object)recebivel.SacadoNome).ToArray() },
                    { "EspecieID", recebiveis.ConvertAll(recebivel => (object)recebivel.EspecieID).ToArray() },
                    { "DataEmissao", recebiveis.ConvertAll(recebivel => (object)recebivel.DataEmissao).ToArray() },
                    { "ValorOriginal", recebiveis.ConvertAll(recebivel => (object)recebivel.ValorOriginal).ToArray() },
                    { "ValorDevido", recebiveis.ConvertAll(recebivel => (object)recebivel.ValorDevido).ToArray() },
                    { "DataBaseCalculo", recebiveis.ConvertAll(recebivel => (object)recebivel.DataBaseCalculo).ToArray() },
                    { "DataInclusao", recebiveis.ConvertAll(recebivel => (object)recebivel.DataInclusao).ToArray() },
                    { "DataVencimento", recebiveis.ConvertAll(recebivel => (object)recebivel.DataVencimento).ToArray() },
                    { "FinanceiroID", recebiveis.ConvertAll(recebivel => (object)recebivel.FinanceiroID).ToArray() },
                    { "LoginID", recebiveis.ConvertAll(recebivel => (object)recebivel.LoginID).ToArray() },
                    { "EstabelecimentoID", recebiveis.ConvertAll(recebivel => (object)recebivel.EstabelecimentoID).ToArray() },
                    { "SituacaoID", recebiveis.ConvertAll(recebivel => (object)recebivel.SituacaoID).ToArray() },
                    { "Observacoes", recebiveis.ConvertAll(recebivel => (object)recebivel.Observacoes).ToArray() },
                    { "ExclusaoMotivo", recebiveis.ConvertAll(recebivel => (object)recebivel.ExclusaoMotivo).ToArray() }
				};

				var salvarArquivo = Tools.GerarNomeArquivo($"Recebiveis_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Recebiveis", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);

				MessageBox.Show("Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

        public void ImportarRecebidos(string arquivoExcel, int estabelecimentoID, int respFinanceiroPessoaID, int loginID, string arquivoExcelRecebiveis, string arquivoExcelRecebidos = "")
        {
			//CRD013 Forma de Pagamento
			var dataHoje = DateTime.Now;
            var indiceLinha = 0;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

			var excelRecebidosDict = ExcelRecebiveisToDictionary(arquivoExcelRecebiveis);

			ISheet sheet;
			try
			{
				IWorkbook workbook = excelHelper.LerExcel(arquivoExcelRecebidos);
				sheet = workbook.GetSheetAt(0);
				excelHelper.InitializeDictionaryRecebidos(sheet);
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelRecebidos}\": {ex.Message}");
			}

			var fluxoCaixas = new List<FluxoCaixa>();
			var recebiveis = new List<Recebivel>();

			try
            {
                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

                    string documento = "";
                    string? pagamento = null, cpf = null, outroSacadoNome = null;
                    int? tipoPagamento = null, consumidorID = null, recebivelID = null;
					string? observacao = null;
                    decimal pagoValor = 0, valor = 0;
                    byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
                    DateTime dataBaixa = DateTime.Now;

                    foreach (var celula in linha.Cells)
                    {
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "DocumentoRef":
										documento = celulaValor;
										break;
									case "ValorPago":
										pagoValor = celulaValor.ArredondarValor();
										break;
									case "ValorOriginal":
										valor = celulaValor.ArredondarValor();
										break;
									case "DataBaixa":
										dataBaixa = celulaValor.ToData();
										break;
									case "ObservaçãoRecebido":
										observacao = celulaValor;
										break;
									//case "TIPO_DOC":
									//tipoPagamento = int.Parse(celulaValor);
									//break;
									case "CPF":
										cpf = celulaValor.ToCPF();
									break;
									case "NOME_GRUPO":
										break;
								}
							}
						}
					}

					//if (tipoPagamento != null && excelFormaPagamentoDict.ContainsKey((int)tipoPagamento))
					//    formaPagamento = (byte)excelFormaPagamentoDict[(int)tipoPagamento];

					var consumidorIDValue = excelHelper.GetConsumidorID(cpf: cpf);

					if (string.IsNullOrEmpty(consumidorIDValue) == false)
					{
						consumidorID = int.Parse(consumidorIDValue);

						if (excelHelper.RecebidoExists((int)consumidorID, pagoValor, dataBaixa) == false)
						{
							if (!string.IsNullOrEmpty(documento) && excelRecebidosDict.ContainsKey(documento))
							{
								recebivelID = int.Parse(excelRecebidosDict[documento][0]);
								consumidorID = int.Parse(excelRecebidosDict[documento][1]);

								recebiveis.Add(new Recebivel()
								{
									ID = int.Parse(excelRecebidosDict[documento][0]),
									DataBaixa = dataBaixa,
									ValorDevido = valor
								});
							}

							var tituloTransacao = TituloTransacoes.Liquidacao;
							if (pagoValor < valor)
								tituloTransacao = TituloTransacoes.PagamentoParcial;

							fluxoCaixas.Add(new FluxoCaixa()
							{
								RecebivelID = recebivelID,
								ConsumidorID = consumidorID,
								SituacaoID = 1,
								PagoMulta = 0,
								PagoJuros = 0,
								PagoDescontos = 0,
								PagoDespesas = 0,
								TipoID = (byte)TransacaoTiposID.Recebimento,
								Data = dataBaixa,
								TransacaoID = (byte)tituloTransacao,
								EspecieID = formaPagamento,
								DataBaseCalculo = dataBaixa,
								DevidoValor = pagoValor,
								PagoValor = pagoValor,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								DataInclusao = dataBaixa,
								FinanceiroID = respFinanceiroPessoaID,
								Observacoes = observacao,
								OutroSacadoNome = outroSacadoNome
							});
						}
						else
							consumidorID = consumidorID;
					}
                }

                indiceLinha = 0;

                var dados = new Dictionary<string, object[]>
                {
                    { "ConsumidorID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ConsumidorID).ToArray() },
                    { "RecebivelID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.RecebivelID).ToArray() },
                    { "SituacaoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.SituacaoID).ToArray() },
                    { "PagoMulta", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoMulta).ToArray() },
                    { "PagoJuros", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoJuros).ToArray() },
                    { "TipoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TipoID).ToArray() },
                    { "Data", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Data).ToArray() },
                    { "TransacaoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TransacaoID).ToArray() },
                    { "EspecieID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EspecieID).ToArray() },
                    { "DataBaseCalculo", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataBaseCalculo).ToArray() },
                    { "DevidoValor", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DevidoValor).ToArray() },
                    { "PagoValor", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoValor).ToArray() },
                    { "EstabelecimentoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EstabelecimentoID).ToArray() },
                    { "LoginID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.LoginID).ToArray() },
                    { "DataInclusao", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataInclusao).ToArray() },
                    { "FinanceiroID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.FinanceiroID).ToArray() },
                    { "Observacoes", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Observacoes).ToArray() },
					{ "OutroSacadoNome", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.OutroSacadoNome).ToArray() }
				};

                var salvarArquivo = Tools.GerarNomeArquivo($"Recebidos_{estabelecimentoID}_OdontoCompany_Migração");
                sqlHelper.GerarSqlInsert("FluxoCaixa", salvarArquivo, dados);
                excelHelper.GravarExcel(salvarArquivo, dados);


				var dadosRecebivel = new Dictionary<string, object[]>
				{
					{ "ID", recebiveis.ConvertAll(recebivel => (object)recebivel.ID).ToArray() },
					{ "DataBaixa", recebiveis.ConvertAll(recebivel => (object)recebivel.DataBaixa).ToArray() },
					{ "ValorDevido", recebiveis.ConvertAll(recebivel => (object)recebivel.ValorDevido).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Recebidos_{estabelecimentoID}_Update_OdontoCompany_Migração");
				sqlHelper.GerarSqlUpdate("Recebiveis", salvarArquivo, dadosRecebivel);

				MessageBox.Show("Sucesso!");
			}
            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }
		public void ImportarPessoas(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			//if (Path.GetFileNameWithoutExtension(arquivoExcel).Contains("CED006"))
				ImportarPessoasDentistas(arquivoExcel, arquivoPessoasAtuais, estabelecimentoID, loginID);
			//else if (Path.GetFileNameWithoutExtension(arquivoExcel).Contains("EMD101"))
				//ImportarPessoasClientes(arquivoExcel, arquivoPessoasAtuais, estabelecimentoID, loginID);
		}

		public void ImportarPessoasClientes(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoPessoasAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoPessoasAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoPessoasAtuais}\": {ex.Message}");
				}


			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var pessoas = new List<Pessoa>();

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
					long? telefonePrinc = null, telefoneAltern = null, telefoneComercial = null, telefoneOutro = null, celular = null;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

					if (indiceLinha == 6255)
						indiceLinha = indiceLinha;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "Código":
										break;
									case "Ativo(S/N)":
										break;
									case "NomeCompleto":
										nomeCompleto = celulaValor.ToNome();
										break;
									case "NomeSocial":
										break;
									case "Apelido":
										apelido = nomeCompleto.GetPrimeirosCaracteres(20).ToNome();
										break;
									case "Documento(CPF,CNPJ,CGC)":
										documento = celulaValor.ToCPF();
										break;
									case "DataCadastro(01/12/2024)":
										dataCadastro = celulaValor.ToData();
										break;
									case "Observações":
										observacao = celulaValor;
										break;
									case "Email":
										email = celulaValor.ToEmail();
										break;
									case "RG":
										rg = celulaValor.GetPrimeirosCaracteres(20);
										break;
									case "Sexo(M/F)":
										sexo = celulaValor.ToSexo("m", "f");
										break;
									case "NascimentoData":
										dataNascimento = celulaValor.ToData();
										break;
									case "NascimentoLocal":
										break;
									case "EstadoCivil(S/C/V)":
										break;
									case "Profissao":
										break;
									case "CargoNaClinica":
										break;
									case "Dentista(S/N)":
										break;
									case "ConselhoCodigo":
										break;
									case "Paciente(S/N)":
										cliente = celulaValor == "S" ? true : false;
										break;
									case "Funcionario(S/N)":
										break;
									case "Fornecedor(S/N)":
										fornecedor = celulaValor == "S" ? true : false;
										break;
									case "TelefonePrincipal":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "Celular":
										celular = celulaValor.ToFone();
										break;
									case "TelefoneAlternativo":
										telefoneAltern = celulaValor.ToFone();
										break;
									case "Logradouro":
										logradouro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "LogradouroNum":
										logradouroNum = celulaValor;
										break;
									case "Complemento":
										break;
									case "Bairro":
										bairro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "Cidade":
										cidade = celulaValor;
										break;
									case "Estado(SP)":
										estado = celulaValor;
										break;
									case "CEP(00000-000)":
										cep = celulaValor.ToNum();
										break;
								}
							}
						}
					}

					pessoaID = indiceLinha;
					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento);

					if (cliente)
					{
						if ((!string.IsNullOrEmpty(nomeCompleto) && string.IsNullOrEmpty(documento))
							|| (!string.IsNullOrEmpty(documento) && documento.IsCPF()))
							if (string.IsNullOrEmpty(pessoaIDValue))
								pessoas.Add(new Pessoa()
								{
									ID = indiceLinha,
									NomeCompleto = nomeCompleto,
									Apelido = apelido,
									CPF = documento,
									DataInclusao = dataCadastro,
									Email = email,
									RG = rg,
									Sexo = sexo,
									NascimentoData = dataNascimento,
									NascimentoLocal = nascimentoLocal,
									ProfissaoOutra = profissaoOutra,
									EstadoCivilID = estadoCivil,
									EstabelecimentoID = estabelecimentoID,
									LoginID = loginID,
									Guid = new Guid(),
									FoneticaApelido = apelido.Fonetizar(),
									FoneticaNomeCompleto = nomeCompleto.Fonetizar()
								});
					}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";


				var pessoasDict = new Dictionary<string, object[]>
				{
					{ "NomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.NomeCompleto).ToArray() },
					{ "Apelido", pessoas.ConvertAll(pessoa => (object)pessoa.Apelido).ToArray() },
					{ "CPF", pessoas.ConvertAll(pessoa => (object)pessoa.CPF).ToArray() },
					{ "DataInclusao", pessoas.ConvertAll(pessoa => (object)pessoa.DataInclusao).ToArray() },
					{ "Email", pessoas.ConvertAll(pessoa => (object)pessoa.Email).ToArray() },
					{ "RG", pessoas.ConvertAll(pessoa => (object)pessoa.RG).ToArray() },
					{ "Sexo", pessoas.ConvertAll(pessoa => (object)pessoa.Sexo).ToArray() },
					{ "NascimentoData", pessoas.ConvertAll(pessoa => (object)pessoa.NascimentoData).ToArray() },
					{ "NascimentoLocal", pessoas.ConvertAll(pessoa => (object)pessoa.NascimentoLocal).ToArray() },
					{ "ProfissaoOutra", pessoas.ConvertAll(pessoa => (object)pessoa.ProfissaoOutra).ToArray() },
					{ "EstadoCivilID", pessoas.ConvertAll(pessoa => (object)pessoa.EstadoCivilID).ToArray() },
					{ "EstabelecimentoID", pessoas.ConvertAll(pessoa => (object)pessoa.EstabelecimentoID).ToArray() },
					{ "LoginID", pessoas.ConvertAll(pessoa => (object)pessoa.LoginID).ToArray() },
					{ "Guid", pessoas.ConvertAll(pessoa => (object)pessoa.Guid).ToArray() },
					{ "FoneticaApelido", pessoas.ConvertAll(pessoa => (object)pessoa.FoneticaApelido).ToArray() },
					{ "FoneticaNomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.FoneticaNomeCompleto).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"PessoasClientes_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Pessoas", salvarArquivo, pessoasDict);
				excelHelper.GravarExcel(salvarArquivo, pessoasDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarEmpresas(string arquivoExcel, string arquivoEmpresasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoEmpresasAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoEmpresasAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoEmpresasAtuais}\": {ex.Message}");
				}


			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var empresas = new List<Empresa>();

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
					long? telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "CLIENTE":
										cliente = celulaValor == "S" ? true : false;
										break;
									case "FORNECEDOR":
										fornecedor = celulaValor == "S" ? true : false;
										break;
									case "NOME":
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										apelido = celulaValor.GetLetras().GetPrimeiroNome().PrimeiraLetraMaiuscula();
										break;
									case "CGC_CPF":
										documento = celulaValor.ToCPF();
										break;
									case "INSC_RG":
										rg = celulaValor.GetPrimeirosCaracteres(20);
										break;
									case "SEXO_M_F":
										sexo = celulaValor.ToSexo("m", "f");
										break;
									case "EMAIL":
										email = celulaValor.ToEmail();
										break;
									case "FONE1":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "FONE2":
										telefoneAltern = celulaValor.ToFone();
										break;
									case "CELULAR":
										celular = celulaValor.ToFone();
										break;
									case "ENDERECO":
										logradouro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "BAIRRO":
										bairro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "NUM_ENDERECO":
										logradouroNum = celulaValor;
										break;
									case "CIDADE":
										cidade = celulaValor;
										break;
									case "ESTADO":
										estado = celulaValor;
										break;
									case "CEP":
										cep = celulaValor.ToNum();
										break;
									case "OBS1":
										observacao = celulaValor;
										break;
									case "NUM_CONVENIO":
										break;
									case "DT_CADASTRO":
										dataCadastro = celulaValor.ToData();
										break;
									case "DT_NASCIMENTO":
										dataNascimento = celulaValor.ToData();
										break;
									case "NUM_FICHA":
										numcadastro = celulaValor;
										break;
								}
							}
						}
					}

					if (string.IsNullOrEmpty(arquivoEmpresasAtuais))
						if (fornecedor && documento.IsCNPJ_CGC())
							empresas.Add(new Empresa()
							{
								Ativo = true,
								CNPJ = documento,
								DataInclusao = dataCadastro,
								LoginID = loginID,
								Marca = "",
								NomeFantasia = "",
								RazaoSocial = "",
								RegimeTribID = 0,
								InscricaoMunicipal = rg,
								EstabelecimentoID = estabelecimentoID,
								Guid = new Guid()
							});
						else
						{
							pessoaID = indiceLinha;
							var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento);
							if (!string.IsNullOrEmpty(pessoaIDValue))
								pessoaID = int.Parse(pessoaIDValue);

							var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento);
							if (!string.IsNullOrEmpty(consumidorIDValue))
								consumidorID = int.Parse(consumidorIDValue);

							if (fornecedor && documento.IsCNPJ_CGC())
								empresas.Add(new Empresa()
								{
									Ativo = true,
									CNPJ = documento,
									DataInclusao = dataCadastro,
									LoginID = loginID,
									Marca = "",
									NomeFantasia = "",
									RazaoSocial = "",
									RegimeTribID = 0,
									InscricaoMunicipal = rg,
									EstabelecimentoID = estabelecimentoID,
									Guid = new Guid()
								});
						}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				var empresasDict = new Dictionary<string, object[]>
				{
					{ "Ativo", empresas.ConvertAll(empresa => (object)empresa.Ativo).ToArray() },
					{ "CNPJ", empresas.ConvertAll(empresa => (object)empresa.CNPJ).ToArray() },
					{ "DataInclusao", empresas.ConvertAll(empresa => (object)empresa.DataInclusao).ToArray() },
					{ "LoginID", empresas.ConvertAll(empresa => (object)empresa.LoginID).ToArray() },
					{ "Marca", empresas.ConvertAll(empresa => (object)empresa.Marca).ToArray() },
					{ "NomeFantasia", empresas.ConvertAll(empresa => (object)empresa.NomeFantasia).ToArray() },
					{ "RazaoSocial", empresas.ConvertAll(empresa => (object)empresa.RazaoSocial).ToArray() },
					{ "RegimeTribID", empresas.ConvertAll(empresa => (object)empresa.RegimeTribID).ToArray() },
					{ "InscricaoMunicipal", empresas.ConvertAll(empresa => (object)empresa.InscricaoMunicipal).ToArray() },
					{ "EstabelecimentoID", empresas.ConvertAll(empresa => (object)empresa.EstabelecimentoID).ToArray() },
					{ "Guid", empresas.ConvertAll(empresa => (object)empresa.Guid).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Empresas_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Empresas", salvarArquivo, empresasDict);
				excelHelper.GravarExcel(salvarArquivo, empresasDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarPessoasDentistas(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();
			bool funcionario = false;

			if (!string.IsNullOrEmpty(arquivoPessoasAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoPessoasAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoPessoasAtuais}\": {ex.Message}");
				}

			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var pessoas = new List<Pessoa>();
				var funcionarios = new List<Funcionario>();
				var enderecos = new List<Endereco>();

				foreach (var linha in excelHelper.linhas)
				{
					bool ativo = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					int? codigo = null;
					long? telefonePrinc = null;
					string? nomeCompleto = null, departamento = null, cro = null, observacao = null, email = null;
					string apelido = "", documento = "";

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									//case "DEPARTAMENTO":
									//	departamento = celulaValor;
									//	break;
									//case "MODIFICADO":
									//	dataCadastro = celulaValor.ToData();
									//	break;

									case "Código":
										codigo = int.Parse(celulaValor);
										break;
									case "Ativo(S/N)":
										ativo = celulaValor == "S" ? true : false;
										break;
									case "NomeCompleto":
										nomeCompleto = celulaValor.ToNome();
										break;
									case "NomeSocial":
										break;
									case "Apelido":
										apelido = celulaValor.GetPrimeirosCaracteres(20).ToNome();
										break;
									case "Documento(CPF,CNPJ,CGC)":
										documento = celulaValor.ToCPF();
										break;
									case "DataCadastro(01/12/2024)":
										dataCadastro = celulaValor.ToData();
										break;
									case "Observações":
										observacao = celulaValor;
										break;
									case "Email":
										email = celulaValor.ToEmail();
										break;
									case "NascimentoData":
										dataNascimento = celulaValor.ToData();
										break;
									case "Funcionario(S/N)":
										funcionario = celulaValor == "S" ? true : false;
										break;
									case "TelefonePrincipal":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "CEP(00000-000)":
										cep = celulaValor.ToNum();
										break;
									case "ConselhoCodigo":
										cro = celulaValor;
										break;
								}
							}
						}
					}

					if (funcionario)
					{
						var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto);

						if (string.IsNullOrEmpty(arquivoPessoasAtuais) 
							|| (!string.IsNullOrEmpty(arquivoPessoasAtuais) && string.IsNullOrEmpty(pessoaIDValue)))
						{
							pessoas.Add(new Pessoa()
							{
								ID = indiceLinha,
								NomeCompleto = nomeCompleto,
								Apelido = apelido,
								CPF = "",
								DataInclusao = dataCadastro,
								Email = email,
								NascimentoData = dataNascimento,
								ProfissaoOutra = departamento,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								Guid = new Guid(),
								FoneticaApelido = apelido.Fonetizar(),
								FoneticaNomeCompleto = nomeCompleto.Fonetizar(),
								Sexo = true,
								ConselhoCodigo = cro
							});
						}
					}
					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				var pessoasDict = new Dictionary<string, object[]>
				{
					//{ "ID", pessoas.ConvertAll(pessoa => (object)pessoa.ID).ToArray() },
					{ "NomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.NomeCompleto).ToArray() },
					{ "Apelido", pessoas.ConvertAll(pessoa => (object)pessoa.Apelido).ToArray() },
					{ "CPF", pessoas.ConvertAll(pessoa => (object)pessoa.CPF).ToArray() },
					{ "DataInclusao", pessoas.ConvertAll(pessoa => (object)pessoa.DataInclusao).ToArray() },
					{ "Email", pessoas.ConvertAll(pessoa => (object)pessoa.Email).ToArray() },
					{ "RG", pessoas.ConvertAll(pessoa => (object)pessoa.RG).ToArray() },
					{ "Sexo", pessoas.ConvertAll(pessoa => (object)pessoa.Sexo).ToArray() },
					{ "NascimentoData", pessoas.ConvertAll(pessoa => (object)pessoa.NascimentoData).ToArray() },
					{ "NascimentoLocal", pessoas.ConvertAll(pessoa => (object)pessoa.NascimentoLocal).ToArray() },
					{ "ProfissaoOutra", pessoas.ConvertAll(pessoa => (object)pessoa.ProfissaoOutra).ToArray() },
					{ "EstadoCivilID", pessoas.ConvertAll(pessoa => (object)pessoa.EstadoCivilID).ToArray() },
					{ "EstabelecimentoID", pessoas.ConvertAll(pessoa => (object)pessoa.EstabelecimentoID).ToArray() },
					{ "LoginID", pessoas.ConvertAll(pessoa => (object)pessoa.LoginID).ToArray() },
					{ "Guid", pessoas.ConvertAll(pessoa => (object)pessoa.Guid).ToArray() },
					{ "FoneticaApelido", pessoas.ConvertAll(pessoa => (object)pessoa.FoneticaApelido).ToArray() },
					{ "FoneticaNomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.FoneticaNomeCompleto).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Pessoas_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Pessoas", salvarArquivo, pessoasDict);
				excelHelper.GravarExcel(salvarArquivo, pessoasDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarFuncionarios(string arquivoExcel, string arquivoFuncionariosAtuais, int estabelecimentoID, int loginID)
        {
			var indiceLinha = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (File.Exists(arquivoExcelCidades))
				try
				{
					var workbookCidades = excelHelper.LerExcel(arquivoExcelCidades);
					var sheetCidades = workbookCidades.GetSheetAt(0);
					excelHelper.InitializeDictionaryCidade(sheetCidades);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelCidades}\": {ex.Message}");
				}

			//if (!string.IsNullOrEmpty(arquivoFuncionariosAtuais))
			//	try
			//	{
			//		var workbook = excelHelper.LerExcel(arquivoFuncionariosAtuais);
			//		var sheet = workbook.GetSheetAt(0);
			//		excelHelper.InitializeDictionary(sheet);
			//	}
			//	catch (Exception ex)
			//	{
			//		throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoFuncionariosAtuais}\": {ex.Message}");
			//	}


			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var funcionarios = new List<Funcionario>();
				var pessoaFones = new List<PessoaFone>();

				foreach (var linha in excelHelper.linhas)
				{
					bool ativo = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					int? codigo = null;
					long? telefonePrinc = null;
					string? nomeCompleto = null, departamento = null, cro = null, observacao = null, email = null;
					string apelido = "";
					bool funcionario = false;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "Código":
										codigo = int.Parse(celulaValor);
										break;
									case "Ativo(S/N)":
										ativo = celulaValor == "S" ? true : false;
										break;
									case "NomeCompleto":
										nomeCompleto = celulaValor.ToNome();
										break;
									case "NomeSocial":
										break;
									case "Apelido":
										apelido = celulaValor.GetPrimeirosCaracteres(20).ToNome();
										break;
									case "DataCadastro(01/12/2024)":
										dataCadastro = celulaValor.ToData();
										break;
									case "Observações":
										observacao = celulaValor;
										break;
									case "Email":
										email = celulaValor.ToEmail();
										break;
									case "NascimentoData":
										dataNascimento = celulaValor.ToData();
										break;
									case "Funcionario(S/N)":
										funcionario = celulaValor == "S" ? true : false;
										break;
									case "TelefonePrincipal":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "CEP(00000-000)":
										cep = celulaValor.ToNum();
										break;
									case "ConselhoCodigo":
										cro = celulaValor;
										break;
								}
							}
						}
					}

					pessoaID = indiceLinha;

					if (funcionario)
					{
						var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto);
						var funcionarioIDValue = excelHelper.GetFuncionarioID(nomeCompleto: nomeCompleto);

						if (string.IsNullOrEmpty(arquivoFuncionariosAtuais) 
							|| (!string.IsNullOrEmpty(arquivoFuncionariosAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && string.IsNullOrEmpty(funcionarioIDValue)))
						{
							pessoaID = int.Parse(pessoaIDValue);

							funcionarios.Add(new Funcionario()
							{
								Ativo = ativo,
								DataInclusao = dataCadastro,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								PessoaID = pessoaID,
								Observacoes = observacao,
								CargoID = (byte)CargosID.Dentista,
								PermissaoCoordenacao = false,
								PermissaoGrupoID = 0,
								PermissaoModuloAdmin = false,
								PermissaoModuloAtendimentos = true,
								PermissaoModuloPacientes = true,
								PermissaoModuloEstoque = true,
								PermissaoModuloFinanceiro = false
							});

							if (telefonePrinc != null)
								if (string.IsNullOrEmpty(arquivoFuncionariosAtuais)
									|| (!string.IsNullOrEmpty(arquivoFuncionariosAtuais) && !excelHelper.PessoaFoneExists(nomeCompleto, telefonePrinc.ToString())))
								{
									pessoaFones.Add(new PessoaFone()
									{
										PessoaID = pessoaID,
										FoneTipoID = (short)FoneTipos.Principal,
										Telefone = (long)telefonePrinc,
										DataInclusao = dataCadastro,
										LoginID = loginID
									});
								}
						}
					}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				var funcionariosDict = new Dictionary<string, object[]>
				{
					{ "CargoID", funcionarios.ConvertAll(funcionario => (object)funcionario.CargoID).ToArray() },
					{ "Ativo", funcionarios.ConvertAll(funcionario => (object)funcionario.Ativo).ToArray() },
					{ "DataInclusao", funcionarios.ConvertAll(funcionario => (object)funcionario.DataInclusao).ToArray() },
					{ "EstabelecimentoID", funcionarios.ConvertAll(funcionario => (object)funcionario.EstabelecimentoID).ToArray() },
					{ "LoginID", funcionarios.ConvertAll(funcionario => (object)funcionario.LoginID).ToArray() },
					{ "PessoaID", funcionarios.ConvertAll(funcionario => (object)funcionario.PessoaID).ToArray() },
					{ "Observacoes", funcionarios.ConvertAll(funcionario => (object)funcionario.Observacoes).ToArray() },
					{ "PermissaoCoordenacao", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoCoordenacao).ToArray() },
					{ "PermissaoGrupoID", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoGrupoID).ToArray() },
					{ "PermissaoModuloAdmin", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloAdmin).ToArray() },
					{ "PermissaoModuloAtendimentos", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloAtendimentos).ToArray() },
					{ "PermissaoModuloPacientes", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloPacientes).ToArray() },
					{ "PermissaoModuloEstoque", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloEstoque).ToArray() },
					{ "PermissaoModuloFinanceiro", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloFinanceiro).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Funcionarios_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Funcionarios", salvarArquivo, funcionariosDict);
				excelHelper.GravarExcel(salvarArquivo, funcionariosDict);


				var pessoaFonesDict = new Dictionary<string, object[]>
				{
					{ "PessoaID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.PessoaID).ToArray() },
					{ "FoneTipoID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.FoneTipoID).ToArray() },
					{ "Telefone", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.Telefone).ToArray() },
					{ "DataInclusao", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.DataInclusao).ToArray() },
					{ "LoginID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"PessoaFones_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("PessoaFones", salvarArquivo, pessoaFonesDict);
				excelHelper.GravarExcel(salvarArquivo, pessoaFonesDict);

				MessageBox.Show("Limpar o Redis" + Environment.NewLine + "redis-cli.exe -h 127.0.0.1 -n 7 del Equipe:" + estabelecimentoID.ToString("D6") + "-Funcionarios", "Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarFornecedores(string arquivoExcel, string arquivoFornecedoresAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoFornecedoresAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoFornecedoresAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoFornecedoresAtuais}\": {ex.Message}");
				}

			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var fornecedores = new List<Fornecedor>();
				var empresas = new List<Empresa>();
				var enderecos = new List<Endereco>();
				var fornecedorFones = new List<FornecedorFone>();

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
					LogradouroTipos logradouroTipo = LogradouroTipos.Outros;
					long? telefonePrinc = null, telefoneAltern = null, telefoneComercial = null, telefoneOutro = null, celular = null;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "CLIENTE":
										cliente = celulaValor == "S" ? true : false;
										break;
									case "FORNECEDOR":
										fornecedor = celulaValor == "S" ? true : false;
										break;
									case "NOME":
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										apelido = celulaValor.GetLetras().GetPrimeiroNome().PrimeiraLetraMaiuscula();
										break;
									case "CGC_CPF":
										documento = celulaValor.ToCPF();
										break;
									case "INSC_RG":
										rg = celulaValor.GetPrimeirosCaracteres(20);
										break;
									case "SEXO_M_F":
										sexo = celulaValor.ToSexo("m", "f");
										break;
									case "EMAIL":
										email = celulaValor.ToEmail();
										break;
									case "FONE1":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "FONE2":
										telefoneAltern = celulaValor.ToFone();
										break;
									case "CELULAR":
										celular = celulaValor.ToFone();
										break;
									case "ENDERECO":
										logradouro = celulaValor.PrimeiraLetraMaiuscula();
										logradouroTipo = logradouro.GetLogradouroTipo();
										if (logradouroTipo != LogradouroTipos.Outros)
											logradouro = logradouro.RemoverPrimeiroNome();
										break;
									case "BAIRRO":
										bairro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "NUM_ENDERECO":
										logradouroNum = celulaValor;
										break;
									case "CIDADE":
										cidade = celulaValor;
										break;
									case "ESTADO":
										estado = celulaValor;
										break;
									case "CEP":
										cep = celulaValor.ToNum();
										break;
									case "OBS1":
										observacao = celulaValor;
										break;
									case "NUM_CONVENIO":
										break;
									case "DT_CADASTRO":
										dataCadastro = celulaValor.ToData();
										break;
									case "DT_NASCIMENTO":
										dataNascimento = celulaValor.ToData();
										break;
									case "NUM_FICHA":
										numcadastro = celulaValor;
										break;
								}
							}
						}
					}

					pessoaID = indiceLinha;
					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento);
					if (!string.IsNullOrEmpty(pessoaIDValue))
						pessoaID = int.Parse(pessoaIDValue);

					var fornecedorIDValue = excelHelper.GetFornecedorID(nomeCompleto: nomeCompleto, cpf: documento);

					if (!string.IsNullOrWhiteSpace(nomeCompleto))
					{
						if (fornecedor && string.IsNullOrEmpty(fornecedorIDValue))
						{
							fornecedores.Add(new Fornecedor()
							{
								Ativo = true,
								DataInclusao = dataCadastro,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								Email = email,
								EmpresaID = indiceLinha,
								Observacoes = observacao
							});

							if (!string.IsNullOrWhiteSpace(cidade) && !string.IsNullOrWhiteSpace(logradouro))
							{
								var cidadeID = excelHelper.GetCidadeID(cidade, estado);
								enderecos.Add(new Endereco()
								{
									Ativo = true,
									Cep = cep,
									CidadeID = cidadeID,
									DataInclusao = dataCadastro,
									EnderecoTipoID = (byte)EnderecoTipos.Residencial,
									Logradouro = logradouro,
									LogradouroNum = logradouroNum,
									Bairro = bairro,
									LogradouroTipoID = (int)logradouroTipo,
									Complemento = complemento,
									ParentID = indiceLinha,
									TableID = 1,
								});
							}

							if (celular != null && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, celular.ToString()))
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Celular,
									Telefone = (long)celular,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefonePrinc != null && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefonePrinc.ToString()))
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Principal,
									Telefone = (long)telefonePrinc,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefoneAltern != null && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefoneAltern.ToString()))
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Alternativo,
									Telefone = (long)telefoneAltern,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefoneComercial != null && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefoneComercial.ToString()))
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Comercial,
									Telefone = (long)telefoneComercial,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefoneOutro != null && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefoneOutro.ToString()))
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Outros,
									Telefone = (long)telefoneOutro,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});
						}
					}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";


				var fornecedorFonesDict = new Dictionary<string, object[]>
				{
					{ "FornecedorID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.FornecedorID).ToArray() },
					{ "FoneTipoID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.FoneTipoID).ToArray() },
					{ "Telefone", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.Telefone).ToArray() },
					{ "DataInclusao", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.DataInclusao).ToArray() },
					{ "LoginID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"FornecedorFones_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoFornecedorFones_Temp", salvarArquivo, fornecedorFonesDict);
				excelHelper.GravarExcel(salvarArquivo, fornecedorFonesDict);


				var enderecosDict = new Dictionary<string, object[]>
				{
					{ "Ativo", enderecos.ConvertAll(endereco => (object)endereco.Ativo).ToArray() },
					{ "Cep", enderecos.ConvertAll(endereco => (object)endereco.Cep).ToArray() },
					{ "CidadeID", enderecos.ConvertAll(endereco => (object)endereco.CidadeID).ToArray() },
					{ "DataInclusao", enderecos.ConvertAll(endereco => (object)endereco.DataInclusao).ToArray() },
					{ "EnderecoTipoID", enderecos.ConvertAll(endereco => (object)endereco.EnderecoTipoID).ToArray() },
					{ "Logradouro", enderecos.ConvertAll(endereco => (object)endereco.Logradouro).ToArray() },
					{ "LogradouroNum", enderecos.ConvertAll(endereco => (object)endereco.LogradouroNum).ToArray() },
					{ "Bairro", enderecos.ConvertAll(endereco => (object)endereco.Bairro).ToArray() },
					{ "LogradouroTipoID", enderecos.ConvertAll(endereco => (object)endereco.LogradouroTipoID).ToArray() },
					{ "Complemento", enderecos.ConvertAll(endereco => (object)endereco.Complemento).ToArray() },
					{ "ParentID", enderecos.ConvertAll(endereco => (object)endereco.ParentID).ToArray() },
					{ "TableID", enderecos.ConvertAll(endereco => (object)endereco.TableID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Enderecos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoEnderecos_Temp", salvarArquivo, enderecosDict);
				excelHelper.GravarExcel(salvarArquivo, enderecosDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarPacientes(string arquivoExcel, string arquivoPacientesAtuais, int estabelecimentoID, int loginID)
        {
			var indiceLinha = 1;
            var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            DateTime dataHoje = DateTime.Now;
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            if (File.Exists(arquivoExcelCidades))
                try
                {
                    var workbookCidades = excelHelper.LerExcel(arquivoExcelCidades);
                    var sheetCidades = workbookCidades.GetSheetAt(0);
                    excelHelper.InitializeDictionaryCidade(sheetCidades);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelCidades}\": {ex.Message}");
                }

			if (!string.IsNullOrEmpty(arquivoPacientesAtuais))
			    try
			    {
				    var workbook = excelHelper.LerExcel(arquivoPacientesAtuais);
				    var sheet = workbook.GetSheetAt(0);
				    excelHelper.InitializeDictionary(sheet);
			    }
			    catch (Exception ex)
			    {
				    throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoPacientesAtuais}\": {ex.Message}");
			    }


			try
            {
                var linhasCount = excelHelper.linhas.Count;
                var consumidores = new List<Consumidor>();
                var consumidoresEnderecos = new List<ConsumidorEndereco>();
                var pessoaFones = new List<PessoaFone>();
				var fornecedores = new List<Fornecedor>();
				var empresas = new List<Empresa>();
                var enderecos = new List<Endereco>();
				var fornecedorFones = new List<FornecedorFone>();

				foreach (var linha in excelHelper.linhas)
                {
                    bool cliente = false, fornecedor = false;
                    DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
                    int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true, ativo = true;
					LogradouroTipos logradouroTipo = LogradouroTipos.Outros;
					long? telefonePrinc = null, telefoneAltern = null, telefoneComercial = null, telefoneOutro = null, celular = null;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

					foreach (var celula in linha.Cells)
                    {
                        if (celula != null)
                        {
                            celulaValor = celula.ToString().Trim().Replace("'", "’");
                            tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
                            colunaLetra = excelHelper.GetColumnLetter(celula);

                            if (!string.IsNullOrWhiteSpace(celulaValor))
                            {
                                switch (tituloColuna)
                                {
                                    case "Paciente(S/N)":
										cliente = celulaValor == "S" ? true : false;
                                        break;
                                    case "Ativo(S/N)":
                                        ativo = celulaValor == "S" ? true : false;
                                        break;
                                    case "NomeCompleto":
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
                                        break;
									case "Apelido":
										apelido = celulaValor.GetLetras().GetPrimeiroNome().PrimeiraLetraMaiuscula();
										break;
									case "Documento(CPF,CNPJ,CGC)":
										documento = celulaValor.ToCPF();
                                        break;
                                    case "RG":
                                        rg = celulaValor.GetPrimeirosCaracteres(20);
                                        break;
                                    case "Sexo(M/F)":
										sexo = celulaValor.ToSexo("m", "f");
                                        break;
                                    case "Email":
                                        email = celulaValor.ToEmail();
                                        break;
                                    case "TelefonePrincipal":
                                        telefonePrinc = celulaValor.ToFone();
                                        break;
                                    case "TelefoneAlternativo":
                                        telefoneAltern = celulaValor.ToFone();
                                        break;
                                    case "Celular":
                                        celular = celulaValor.ToFone();
                                        break;
                                    case "Logradouro":
                                        logradouro = celulaValor.PrimeiraLetraMaiuscula();
										logradouroTipo = logradouro.GetLogradouroTipo();
										if (logradouroTipo != LogradouroTipos.Outros)
											logradouro = logradouro.RemoverPrimeiroNome();
										break;
                                    case "Bairro":
                                        bairro = celulaValor.PrimeiraLetraMaiuscula();
                                        break;
                                    case "LogradouroNum":
                                        logradouroNum = celulaValor;
                                        break;
                                    case "Cidade":
                                        cidade = celulaValor;
                                        break;
                                    case "Estado(SP)":
										estado = celulaValor;
                                        break;
                                    case "CEP(00000-000)":
										cep = celulaValor.ToNum();
                                        break;
                                    case "Observações":
                                        observacao = celulaValor;
										break;
                                    case "NUM_CONVENIO":
                                        break;
                                    case "DT_CADASTRO":
                                        dataCadastro = celulaValor.ToData();
                                        break;
                                    case "NascimentoData":
                                        dataNascimento = celulaValor.ToData();
                                        break;
									case "Código":
										numcadastro = celulaValor;
										break;
								}
                            }
                        }
                    }

                    pessoaID = indiceLinha;

					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento);
					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento);

					if (!string.IsNullOrWhiteSpace(nomeCompleto))
                    {
                        if (cliente)
                        {
							if (string.IsNullOrEmpty(pessoaIDValue) == false)
							{
								pessoaID = int.Parse(pessoaIDValue);

								if (string.IsNullOrEmpty(consumidorIDValue))
								{
									consumidores.Add(new Consumidor()
									{
										Ativo = ativo,
										DataInclusao = dataCadastro,
										EstabelecimentoID = estabelecimentoID,
										LGPDSituacaoID = 0,
										LoginID = loginID,
										PessoaID = pessoaID,
										CodigoAntigo = numcadastro,
										Observacoes = observacao
									});
								}
								else
								{
									var cidadeID = excelHelper.GetCidadeID(cidade, estado);

									if (cidadeID > 0)
										if (excelHelper.ConsumidorEnderecoExists(pessoaID, cep) == false)
											consumidoresEnderecos.Add(new ConsumidorEndereco()
											{
												Ativo = ativo,
												ConsumidorID = int.Parse(consumidorIDValue),
												EnderecoTipoID = (short)EnderecoTipos.Residencial,
												LogradouroTipoID = (int)logradouroTipo,
												Logradouro = logradouro,
												CidadeID = cidadeID,
												Cep = cep,
												DataInclusao = dataCadastro,
												Bairro = bairro,
												LogradouroNum = logradouroNum,
												Complemento = complemento,
												LoginID = loginID
											});
								}


								if (celular != null)
									if (excelHelper.PessoaFoneExists(pessoaID, celular.ToString()) == false)
										pessoaFones.Add(new PessoaFone()
										{
											PessoaID = pessoaID,
											FoneTipoID = (short)FoneTipos.Celular,
											Telefone = (long)celular,
											DataInclusao = dataCadastro,
											LoginID = loginID
										});

								if (telefonePrinc != null)
									if (excelHelper.PessoaFoneExists(pessoaID, telefonePrinc.ToString()) == false)
										pessoaFones.Add(new PessoaFone()
										{
											PessoaID = pessoaID,
											FoneTipoID = (short)FoneTipos.Principal,
											Telefone = (long)telefonePrinc,
											DataInclusao = dataCadastro,
											LoginID = loginID
										});

								if (telefoneAltern != null)
									if (excelHelper.PessoaFoneExists(pessoaID, telefoneAltern.ToString()) == false)
										pessoaFones.Add(new PessoaFone()
										{
											PessoaID = pessoaID,
											FoneTipoID = (short)FoneTipos.Alternativo,
											Telefone = (long)telefoneAltern,
											DataInclusao = dataCadastro,
											LoginID = loginID
										});

								if (telefoneComercial != null)
									if (excelHelper.PessoaFoneExists(pessoaID, telefoneComercial.ToString()) == false)
										pessoaFones.Add(new PessoaFone()
										{
											PessoaID = pessoaID,
											FoneTipoID = (short)FoneTipos.Comercial,
											Telefone = (long)telefoneComercial,
											DataInclusao = dataCadastro,
											LoginID = loginID
										});

								if (telefoneOutro != null)
									if (excelHelper.PessoaFoneExists(pessoaID, telefoneOutro.ToString()) == false)
										pessoaFones.Add(new PessoaFone()
										{
											PessoaID = pessoaID,
											FoneTipoID = (short)FoneTipos.Outros,
											Telefone = (long)telefoneOutro,
											DataInclusao = dataCadastro,
											LoginID = loginID
										});
							}
                        }
                    }

					indiceLinha++;
				}

                indiceLinha = 0;
                var salvarArquivo = "";


				var consumidoresDict = new Dictionary<string, object[]>
				{
					{ "Ativo", consumidores.ConvertAll(consumidor => (object)consumidor.Ativo).ToArray() },
                    { "DataInclusao", consumidores.ConvertAll(consumidor => (object)consumidor.DataInclusao).ToArray() },
                    { "EstabelecimentoID", consumidores.ConvertAll(consumidor => (object)consumidor.EstabelecimentoID).ToArray() },
                    { "LGPDSituacaoID", consumidores.ConvertAll(consumidor => (object)consumidor.LGPDSituacaoID).ToArray() },
                    { "LoginID", consumidores.ConvertAll(consumidor => (object)consumidor.LoginID).ToArray() },
                    { "PessoaID", consumidores.ConvertAll(consumidor => (object)consumidor.PessoaID).ToArray() },
                    { "CodigoAntigo", consumidores.ConvertAll(consumidor => (object)consumidor.CodigoAntigo).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Consumidores_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Consumidores", salvarArquivo, consumidoresDict);
				excelHelper.GravarExcel(salvarArquivo, consumidoresDict);


				var consumidoresEnderecosDict = new Dictionary<string, object[]>
				{
					{ "LoginID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.LoginID).ToArray() },
					{ "Ativo", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Ativo).ToArray() },
                    { "ConsumidorID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.ConsumidorID).ToArray() },
                    { "EnderecoTipoID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.EnderecoTipoID).ToArray() },
                    { "LogradouroTipoID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.LogradouroTipoID).ToArray() },
                    { "Logradouro", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Logradouro).ToArray() },
                    { "CidadeID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.CidadeID).ToArray() },
                    { "Cep", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Cep).ToArray() },
                    { "DataInclusao", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.DataInclusao).ToArray() },
                    { "Bairro", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Bairro).ToArray() },
                    { "LogradouroNum", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.LogradouroNum).ToArray() },
                    { "Complemento", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Complemento).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"ConsumidorEnderecos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("ConsumidorEnderecos", salvarArquivo, consumidoresEnderecosDict);
				excelHelper.GravarExcel(salvarArquivo, consumidoresEnderecosDict);
				

				var pessoaFonesDict = new Dictionary<string, object[]>
				{
					{ "PessoaID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.PessoaID).ToArray() },
					{ "FoneTipoID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.FoneTipoID).ToArray() },
					{ "Telefone", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.Telefone).ToArray() },
					{ "DataInclusao", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.DataInclusao).ToArray() },
					{ "LoginID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"PessoaFones_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("PessoaFones", salvarArquivo, pessoaFonesDict);
				excelHelper.GravarExcel(salvarArquivo, pessoaFonesDict);

				MessageBox.Show("Atualize a Base de Pesquisa em: Dados da Clínica => Licença de Uso", "Sucesso!");
			}

            catch (Exception error)
            {
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

		public static Dictionary<string, string[]> ExcelRecebiveisToDictionary(string arquivoExcel)
		{
			var dataDictionary = new Dictionary<string, string[]>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			if (!excelHelper.cabecalhos.Contains("ExclusaoMotivo") && !excelHelper.cabecalhos.Contains("ID") && !excelHelper.cabecalhos.Contains("ConsumidorID"))
				throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas ID, ConsumidorID e/ou ExclusaoMotivo");

			try
			{
                foreach (var linha in excelHelper.linhas)
                {
                    string documento = "", recebivelID = "", consumidorID = "";

                    foreach (var celula in linha.Cells)
                    {
                        var celulaValor = celula.ToString().Trim();
                        var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];							

						switch (tituloColuna)
                        {
                            case "ExclusaoMotivo":
                                documento = celulaValor;
								break;
							case "ID":
								recebivelID = celulaValor;
								break;
							case "ConsumidorID":
								consumidorID = celulaValor;
								break;
							case "CPF":
								consumidorID = celulaValor;
								break;
						}
                    }

                    if (!string.IsNullOrEmpty(documento) && !string.IsNullOrEmpty(recebivelID) && !string.IsNullOrEmpty(consumidorID) && !dataDictionary.ContainsKey(documento))
					    dataDictionary.Add(documento, [recebivelID, consumidorID]);
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			return dataDictionary;
		}

		public static Dictionary<int, TitulosEspeciesID> ExcelFormaPagamentoToDictionary(string arquivoExcel)
        {
			var dataDictionary = new Dictionary<int, TitulosEspeciesID>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			if (!excelHelper.cabecalhos.Contains("CODIGO") && !excelHelper.cabecalhos.Contains("NOME"))
				throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas CODIGO e/ou NOME");

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					int? codigo = null;
					string formaPagamento = "";
					TitulosEspeciesID titulosEspeciesID = TitulosEspeciesID.DepositoEmConta;

					foreach (var celula in linha.Cells)
					{
						var celulaValor = celula.ToString().Trim();
						var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];

						switch (tituloColuna)
						{
							case "CODIGO":
								codigo = int.Parse(celulaValor);
								break;
							case "NOME":
								formaPagamento = celulaValor;
								break;
						}
					}

					if (codigo != null && !string.IsNullOrEmpty(formaPagamento))
                    {
                        if (formaPagamento.ToLower().Contains("dinheiro"))
							titulosEspeciesID = TitulosEspeciesID.Dinheiro;
                        else if (formaPagamento.ToLower().Contains("cheque"))
							titulosEspeciesID = TitulosEspeciesID.Cheque;
						else if (formaPagamento.ToLower().Contains("master card") || formaPagamento.ToLower().Contains("visa") || formaPagamento.ToLower().Contains("elo")
                            || formaPagamento.ToLower().Contains("american express") || formaPagamento.ToLower().Contains("hipercard"))
							titulosEspeciesID = TitulosEspeciesID.CartaoCredito;
						else if (formaPagamento.ToLower().Contains("debito"))
							titulosEspeciesID = TitulosEspeciesID.CartaoDebito;
					}

					dataDictionary.Add((int)codigo, titulosEspeciesID);
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			return dataDictionary;
		}

		public static Dictionary<int, ProcedimentosCategoriasID> GruposProcedimentosToDictionary(string arquivoExcel)
        {
			var dataDictionary = new Dictionary<int, ProcedimentosCategoriasID>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			if (!excelHelper.cabecalhos.Contains("CODIGO") && !excelHelper.cabecalhos.Contains("NOME"))
				throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas CODIGO e/ou NOME");

			var grupoCategoriaDict = new Dictionary<string, ProcedimentosCategoriasID>()
            {
				{ "CIRURGIA", ProcedimentosCategoriasID.Cirurgia },
                { "ENDODONTIA", ProcedimentosCategoriasID.Endodontia },
                { "PERIODONTIA", ProcedimentosCategoriasID.Periodontia },
                { "PROTESE", ProcedimentosCategoriasID.Prótese },
                { "CLINICO", ProcedimentosCategoriasID.Outros },
                { "MANUTENCAO", ProcedimentosCategoriasID.Ortodontia },
                { "ORTODONTIA", ProcedimentosCategoriasID.Ortodontia },
                { "PREVENCAO", ProcedimentosCategoriasID.Prevenção },
                { "OROFACIAL", ProcedimentosCategoriasID.Orofacial }
			};

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
                    int? codigo = null;
                    string procedimentoNome = "";

					foreach (var celula in linha.Cells)
					{
						var celulaValor = celula.ToString().Trim();
						var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];

						switch (tituloColuna)
						{
							case "CODIGO":
								codigo = int.Parse(celulaValor);
								break;
							case "NOME":
								procedimentoNome = celulaValor;
								break;
						}
					}

					if (codigo != null && !string.IsNullOrEmpty(procedimentoNome))
						if (grupoCategoriaDict.ContainsKey(procedimentoNome))
						    dataDictionary.Add((int)codigo, grupoCategoriaDict[procedimentoNome]);
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			return dataDictionary;
		}
	}
}
