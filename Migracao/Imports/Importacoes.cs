﻿using ExcelDataReader.Log.Logger;
using Migracao.Models;
using Migracao.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace Migracao.Imports
{
	internal class Importacoes
	{
		string arquivoExcelCidades = "Files\\EnderecosCidades.xlsx";

		public Importacoes() {
			//ExcelHelper excelHelper = new ExcelHelper();

			//if (File.Exists(arquivoExcelCidades))
			//	try
			//	{
			//		var workbookCidades = excelHelper.LerExcel(arquivoExcelCidades);
			//		var sheetCidades = workbookCidades.GetSheetAt(0);
			//		excelHelper.InitializeDictionaryCidade(sheetCidades);
			//	}
			//	catch (Exception ex)
			//	{
			//		throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelCidades}\": {ex.Message}");
			//	}
		}

		public void Atendimentos(string filePath)
		{
			// Carrega o arquivo Excel
			IWorkbook workbook;
			using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				if (Path.GetExtension(filePath) == ".xls")
				{
					workbook = new HSSFWorkbook(file);
				}
				else
				{
					workbook = new XSSFWorkbook(file);
				}
			}

			// Lê a primeira planilha
			ISheet sheet = workbook.GetSheetAt(0);

			// Cria um DataTable para armazenar os dados
			DataTable dataTable = new DataTable();
			dataTable.Columns.Add("LoginID", typeof(int));
			dataTable.Columns.Add("EstabelecimentoID", typeof(int));
			dataTable.Columns.Add("AtendeTipoID", typeof(int));
			dataTable.Columns.Add("AtendimentoTipoCustomID", typeof(int));
			dataTable.Columns.Add("DataChegada", typeof(DateTime));
			dataTable.Columns.Add("DataInicio", typeof(DateTime));
			dataTable.Columns.Add("DataTermino", typeof(DateTime));
			dataTable.Columns.Add("DataCancelamento", typeof(DateTime));
			dataTable.Columns.Add("ConsumidorID", typeof(int));
			dataTable.Columns.Add("AtendimentoValor", typeof(decimal));
			dataTable.Columns.Add("SecretariaID", typeof(int));
			dataTable.Columns.Add("FuncionarioID", typeof(int));
			dataTable.Columns.Add("SalaID", typeof(int));
			dataTable.Columns.Add("ConvenioID", typeof(int));
			dataTable.Columns.Add("TempoAtraso", typeof(TimeSpan));
			dataTable.Columns.Add("TempoSalaEspera", typeof(TimeSpan));
			dataTable.Columns.Add("TempoAtendimento", typeof(TimeSpan));
			dataTable.Columns.Add("Observacoes", typeof(string));
			dataTable.Columns.Add("DataInclusao", typeof(DateTime));
			dataTable.Columns.Add("DataUltAlteracao", typeof(DateTime));
			dataTable.Columns.Add("AtendimentoIndex", typeof(int));
			dataTable.Columns.Add("EncaminhadoPorMedicoPessoaID", typeof(int));
			dataTable.Columns.Add("DiagnosticoID", typeof(int));

			// Lê os dados das linhas, ignorando o cabeçalho
			for (var row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
				{
					DataRow dataRow = dataTable.NewRow();

					dataRow["LoginID"] = sheet.GetRow(row).GetCell(0) == null || string.IsNullOrEmpty(sheet.GetRow(row).GetCell(0).ToString())
						? DBNull.Value : Convert.ToInt32(sheet.GetRow(row).GetCell(0).NumericCellValue);
					dataRow["EstabelecimentoID"] = sheet.GetRow(row).GetCell(1) == null ? DBNull.Value : Convert.ToInt32(sheet.GetRow(row).GetCell(1).NumericCellValue);
					dataRow["AtendeTipoID"] = sheet.GetRow(row).GetCell(2) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(2));
					dataRow["AtendimentoTipoCustomID"] = sheet.GetRow(row).GetCell(3) == null || string.IsNullOrEmpty(sheet.GetRow(row).GetCell(3).ToString())
						? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(3));
					dataRow["DataChegada"] = sheet.GetRow(row).GetCell(4) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(4));
					dataRow["DataInicio"] = sheet.GetRow(row).GetCell(5) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(5));
					dataRow["DataTermino"] = sheet.GetRow(row).GetCell(6) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(6));
					dataRow["DataCancelamento"] = sheet.GetRow(row).GetCell(7) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(7));
					dataRow["ConsumidorID"] = sheet.GetRow(row).GetCell(8) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(8));
					dataRow["AtendimentoValor"] = sheet.GetRow(row).GetCell(9) == null ? DBNull.Value : Tools.GetDecimalValueFromCell(sheet.GetRow(row).GetCell(9));
					dataRow["SecretariaID"] = sheet.GetRow(row).GetCell(10) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(10));
					dataRow["FuncionarioID"] = sheet.GetRow(row).GetCell(11) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(11));
					dataRow["SalaID"] = sheet.GetRow(row).GetCell(12) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(12));
					dataRow["ConvenioID"] = sheet.GetRow(row).GetCell(13) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(13));
					dataRow["TempoAtraso"] = sheet.GetRow(row).GetCell(14) == null ? DBNull.Value : Tools.GetTimeSpanValueFromCell(sheet.GetRow(row).GetCell(14));
					dataRow["TempoSalaEspera"] = sheet.GetRow(row).GetCell(15) == null ? DBNull.Value : Tools.GetTimeSpanValueFromCell(sheet.GetRow(row).GetCell(15));
					dataRow["TempoAtendimento"] = sheet.GetRow(row).GetCell(16) == null ? DBNull.Value : Tools.GetTimeSpanValueFromCell(sheet.GetRow(row).GetCell(16));
					dataRow["Observacoes"] = sheet.GetRow(row).GetCell(17) == null ? DBNull.Value : sheet.GetRow(row).GetCell(17).StringCellValue;
					dataRow["DataInclusao"] = sheet.GetRow(row).GetCell(18) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(18));
					dataRow["DataUltAlteracao"] = sheet.GetRow(row).GetCell(19) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(19));
					dataRow["AtendimentoIndex"] = sheet.GetRow(row).GetCell(20) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(20));
					dataRow["EncaminhadoPorMedicoPessoaID"] = sheet.GetRow(row).GetCell(21) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(21));
					dataRow["DiagnosticoID"] = sheet.GetRow(row).GetCell(22) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(22));

					dataTable.Rows.Add(dataRow);
				}
			}

			// Gera o script SQL de inserção
			//string sql = GerarSqlInsert(dataTable);

			//File.WriteAllText("asdf.sql", sql);
		}



		public void ImportarConsumidoresPessoas(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();
			List<string> linhasSql = new();

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

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime? dataNascimento = null;
					DateTime dataCadastro = dataHoje;
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

					Pessoa pessoa = null;
					Consumidor consumidor = null;
					ConsumidorEndereco consumidorEndereco = null;
					PessoaFone pessoaFone = null;
					pessoaID = indiceLinha;
					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento, nascimentoData: dataNascimento);

					if (cliente)
					{
						if ((!string.IsNullOrEmpty(nomeCompleto) && string.IsNullOrEmpty(documento))
							|| (!string.IsNullOrEmpty(documento) && documento.IsCPF()))
							if (string.IsNullOrEmpty(pessoaIDValue))
								pessoa = new Pessoa()
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
								};

						consumidor = new Consumidor()
						{
							Ativo = true,
							DataInclusao = dataCadastro,
							EstabelecimentoID = estabelecimentoID,
							LGPDSituacaoID = 0,
							LoginID = loginID,
							PessoaID = pessoaID,
							CodigoAntigo = numcadastro,
							Observacoes = observacao
						};
					}
					else
					{
						var cidadeID = excelHelper.GetCidadeID(cidade, estado);

						if (cidadeID > 0)
							if (excelHelper.ConsumidorEnderecoExists(pessoaID, cep) == false)
								consumidorEndereco = new ConsumidorEndereco()
								{
									Ativo = true,
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
								};
					}


					if (celular != null)
						if (excelHelper.PessoaFoneExists(pessoaID, celular.ToString()) == false)
							pessoaFone = new PessoaFone()
							{
								PessoaID = pessoaID,
								FoneTipoID = (short)FoneTipos.Celular,
								Telefone = (long)celular,
								DataInclusao = dataCadastro,
								LoginID = loginID
							};

					if (telefonePrinc != null)
						if (excelHelper.PessoaFoneExists(pessoaID, telefonePrinc.ToString()) == false)
							pessoaFone = new PessoaFone()
							{
								PessoaID = pessoaID,
								FoneTipoID = (short)FoneTipos.Principal,
								Telefone = (long)telefonePrinc,
								DataInclusao = dataCadastro,
								LoginID = loginID
							});

					if (telefoneAltern != null)
						if (excelHelper.PessoaFoneExists(pessoaID, telefoneAltern.ToString()) == false)
							pessoaFone = new PessoaFone()
							{
								PessoaID = pessoaID,
								FoneTipoID = (short)FoneTipos.Alternativo,
								Telefone = (long)telefoneAltern,
								DataInclusao = dataCadastro,
								LoginID = loginID
							};

					if (telefoneComercial != null)
						if (excelHelper.PessoaFoneExists(pessoaID, telefoneComercial.ToString()) == false)
							pessoaFone = new PessoaFone()
							{
								PessoaID = pessoaID,
								FoneTipoID = (short)FoneTipos.Comercial,
								Telefone = (long)telefoneComercial,
								DataInclusao = dataCadastro,
								LoginID = loginID
							};

					if (telefoneOutro != null)
						if (excelHelper.PessoaFoneExists(pessoaID, telefoneOutro.ToString()) == false)
							pessoaFone = new PessoaFone()
							{
								PessoaID = pessoaID,
								FoneTipoID = (short)FoneTipos.Outros,
								Telefone = (long)telefoneOutro,
								DataInclusao = dataCadastro,
								LoginID = loginID
							};

					indiceLinha++;


					var pessoaDict = new Dictionary<string, object>
					{
						{ "NomeCompleto", pessoa.NomeCompleto },
						{ "Apelido", pessoa.Apelido },
						{ "CPF", pessoa.CPF },
						{ "DataInclusao", pessoa.DataInclusao },
						{ "Email", pessoa.Email },
						{ "RG", pessoa.RG },
						{ "Sexo", pessoa.Sexo },
						{ "NascimentoData", pessoa.NascimentoData },
						{ "NascimentoLocal", pessoa.NascimentoLocal },
						{ "ProfissaoOutra", pessoa.ProfissaoOutra },
						{ "EstadoCivilID", pessoa.EstadoCivilID },
						{ "EstabelecimentoID", pessoa.EstabelecimentoID },
						{ "LoginID", pessoa.LoginID },
						{ "Guid", pessoa.Guid },
						{ "FoneticaApelido", pessoa.FoneticaApelido },
						{ "FoneticaNomeCompleto", pessoa.FoneticaNomeCompleto }
					};

					var consumidorDict = new Dictionary<string, object>
					{
						{ "Ativo", consumidor.Ativo },
						{ "DataInclusao", consumidor.DataInclusao },
						{ "EstabelecimentoID", consumidor.EstabelecimentoID },
						{ "LGPDSituacaoID", consumidor.LGPDSituacaoID },
						{ "LoginID", consumidor.LoginID },
						{ "PessoaID", consumidor.PessoaID },
						{ "CodigoAntigo", consumidor.CodigoAntigo }
					};


					var consumidorEnderecoDict = new Dictionary<string, object>
					{
						{ "LoginID", consumidorEndereco.LoginID },
						{ "Ativo", consumidorEndereco.Ativo },
						{ "ConsumidorID", consumidorEndereco.ConsumidorID },
						{ "EnderecoTipoID", consumidorEndereco.EnderecoTipoID },
						{ "LogradouroTipoID", consumidorEndereco.LogradouroTipoID },
						{ "Logradouro", consumidorEndereco.Logradouro },
						{ "CidadeID", consumidorEndereco.CidadeID },
						{ "Cep", consumidorEndereco.Cep },
						{ "DataInclusao", consumidorEndereco.DataInclusao },
						{ "Bairro", consumidorEndereco.Bairro },
						{ "LogradouroNum", consumidorEndereco.LogradouroNum },
						{ "Complemento", consumidorEndereco.Complemento }
					};


					var pessoaFoneDict = new Dictionary<string, object>
					{
						{ "PessoaID", pessoaFone.PessoaID },
						{ "FoneTipoID", pessoaFone.FoneTipoID },
						{ "Telefone", pessoaFone.Telefone },
						{ "DataInclusao", pessoaFone.DataInclusao },
						{ "LoginID", pessoaFone.LoginID }
					};

					linhasSql.Add(sqlHelper.GerarSqlInsert(1, pessoaDict, consumidorDict));
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				salvarArquivo = Tools.GerarNomeArquivo($"PessoasClientes_{estabelecimentoID}_OdontoCompany_Migração");
				File.WriteAllLines(salvarArquivo + ".sql", linhasSql);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
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


					if (!excelHelper.ExisteTexto(sheet, "Titulo", titulo))
					{

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
			ImportarPessoasDentistas(arquivoExcel, arquivoPessoasAtuais, estabelecimentoID, loginID);
			ImportarPessoasClientes(arquivoExcel, arquivoPessoasAtuais, estabelecimentoID, loginID);
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
					DateTime? dataNascimento = null;
					DateTime dataCadastro = dataHoje;
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
					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento, nascimentoData: dataNascimento);

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
					DateTime? dataNascimento = null;
					DateTime dataCadastro = dataHoje;
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

				salvarArquivo = Tools.GerarNomeArquivo($"PessoasDentistas_{estabelecimentoID}_OdontoCompany_Migração");
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

					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento, nascimentoData: dataNascimento);
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