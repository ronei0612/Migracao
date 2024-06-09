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

		public Importacoes()
		{
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



		public void ImportarPessoas(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();
			List<string> linhasSql = new();

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

			try
			{
				var linhasCount = excelHelper.linhas.Count;

				foreach (var linha in excelHelper.linhas)
				{
					bool iscliente = false, isfornecedor = false, isfuncionario = false, sexo = true, ativo = true, isdentista = false;
					DateTime? dataNascimento = null;
					DateTime dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					long? telefonePrinc = null, telefoneAltern = null, telefoneComercial = null, telefoneOutro = null, celular = null;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "", codigo = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;
					LogradouroTipos logradouroTipo = LogradouroTipos.Outros;


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
										codigo = celulaValor;
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
										sexo = celulaValor == "S" ? true : celulaValor == "N" ? false : celulaValor.ToSexo("m", "f");
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
										iscliente = celulaValor == "S" ? true : false;
										break;
									case "Funcionario(S/N)":
										isfuncionario = celulaValor == "S" ? true : false;
										break;
									case "Fornecedor(S/N)":
										isfornecedor = celulaValor == "S" ? true : false;
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
										logradouroTipo = logradouro.GetLogradouroTipo();
										if (logradouroTipo != LogradouroTipos.Outros)
											logradouro = logradouro.RemoverPrimeiroNome();
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

					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento, nascimentoData: dataNascimento);
					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento, codigo: codigo);
					var funcionarioIDValue = excelHelper.GetFuncionarioID(nomeCompleto: nomeCompleto);
					var cidadeID = excelHelper.GetCidadeID(cidade, estado);
					int pessoaID = 0;
					int consumidorID = 0;
					//int funcionarioID = 0;

					if (!string.IsNullOrEmpty(pessoaIDValue))
						pessoaID = int.Parse(pessoaIDValue);

					if (!string.IsNullOrEmpty(consumidorIDValue))
						consumidorID = int.Parse(consumidorIDValue);

					//if (!string.IsNullOrEmpty(funcionarioIDValue))
					//	funcionarioID = int.Parse(funcionarioIDValue);

					Pessoa pessoa = null;
					Consumidor consumidor = null;
					ConsumidorEndereco consumidorEndereco = null;
					Funcionario funcionario = null;
					Endereco endereco = null;
					List<PessoaFone> pessoaFones = new();

					if (!isfornecedor)
					{
						if (string.IsNullOrEmpty(consumidorIDValue) && string.IsNullOrEmpty(pessoaIDValue) && !string.IsNullOrEmpty(nomeCompleto)
							&& ((!string.IsNullOrEmpty(documento) && documento.IsCPF())
								|| string.IsNullOrEmpty(documento)))
						{
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

						if (isdentista && string.IsNullOrEmpty(funcionarioIDValue))
						{
							funcionario = new Funcionario()
							{
								Ativo = ativo,
								DataInclusao = dataCadastro,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								Observacoes = observacao,
								CargoID = (byte)CargosID.Dentista,
								PermissaoCoordenacao = false,
								PermissaoGrupoID = 0,
								PermissaoModuloAdmin = false,
								PermissaoModuloAtendimentos = true,
								PermissaoModuloPacientes = true,
								PermissaoModuloEstoque = true,
								PermissaoModuloFinanceiro = false
							};

							endereco = new Endereco()
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
								TableID = 1
							};
						}

						else if (iscliente)
						{
							if (string.IsNullOrEmpty(consumidorIDValue))
								consumidor = new Consumidor()
								{
									Ativo = true,
									DataInclusao = dataCadastro,
									EstabelecimentoID = estabelecimentoID,
									LGPDSituacaoID = 0,
									LoginID = loginID,
									CodigoAntigo = numcadastro,
									Observacoes = observacao
								};

							if (cidadeID > 0 && !excelHelper.ConsumidorEnderecoExists(nomeCompleto, cep))
								consumidorEndereco = new ConsumidorEndereco()
								{
									Ativo = true,
									ConsumidorID = consumidorID,
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

						Dictionary<string, object>? pessoaDict = null;
						Dictionary<string, object>? consumidorDict = null;
						Dictionary<string, object>? consumidorEnderecoDict = null;
						Dictionary<string, object[]>? pessoaFonesDict = null;
						Dictionary<string, object>? funcionarioDict = null;
						Dictionary<string, object>? enderecoDict = null;

						if (pessoa != null)
							pessoaDict = new Dictionary<string, object>
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

						if (consumidor != null)
							consumidorDict = new Dictionary<string, object>
						{
							{ "Ativo", consumidor.Ativo },
							{ "DataInclusao", consumidor.DataInclusao },
							{ "EstabelecimentoID", consumidor.EstabelecimentoID },
							{ "LGPDSituacaoID", consumidor.LGPDSituacaoID },
							{ "LoginID", consumidor.LoginID },
							{ "CodigoAntigo", consumidor.CodigoAntigo }
						};

						if (consumidorEndereco != null)
							consumidorEnderecoDict = new Dictionary<string, object>
						{
							{ "LoginID", consumidorEndereco.LoginID },
							{ "Ativo", consumidorEndereco.Ativo },
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

						if (pessoaFones.Count > 0)
							pessoaFonesDict = new Dictionary<string, object[]>
						{
							{ "FoneTipoID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.FoneTipoID).ToArray() },
							{ "Telefone", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.Telefone).ToArray() },
							{ "DataInclusao", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.DataInclusao).ToArray() },
							{ "LoginID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.LoginID).ToArray() }
						};

						if (funcionario != null)
							funcionarioDict = new Dictionary<string, object>
						{
							{ "CargoID", funcionario.CargoID },
							{ "Ativo", funcionario.Ativo },
							{ "DataInclusao", funcionario.DataInclusao },
							{ "EstabelecimentoID", funcionario.EstabelecimentoID },
							{ "LoginID", funcionario.LoginID },
							{ "Observacoes", funcionario.Observacoes },
							{ "PermissaoCoordenacao", funcionario.PermissaoCoordenacao },
							{ "PermissaoGrupoID", funcionario.PermissaoGrupoID },
							{ "PermissaoModuloAdmin", funcionario.PermissaoModuloAdmin },
							{ "PermissaoModuloAtendimentos", funcionario.PermissaoModuloAtendimentos },
							{ "PermissaoModuloPacientes", funcionario.PermissaoModuloPacientes },
							{ "PermissaoModuloEstoque", funcionario.PermissaoModuloEstoque },
							{ "PermissaoModuloFinanceiro", funcionario.PermissaoModuloFinanceiro }
						};

						if (endereco != null)
							enderecoDict = new Dictionary<string, object>
						{
							{ "Ativo", endereco.Ativo },
							{ "Cep", endereco.Cep },
							{ "CidadeID", endereco.CidadeID },
							{ "DataInclusao", endereco.DataInclusao },
							{ "EnderecoTipoID", endereco.EnderecoTipoID },
							{ "Logradouro", endereco.Logradouro },
							{ "LogradouroNum", endereco.LogradouroNum },
							{ "Bairro", endereco.Bairro },
							{ "LogradouroTipoID", endereco.LogradouroTipoID },
							{ "Complemento", endereco.Complemento },
							{ "TableID", endereco.TableID }
						};

						if (pessoaDict != null || pessoaFonesDict != null || consumidorDict != null || consumidorEnderecoDict != null || funcionarioDict != null || enderecoDict != null)
							linhasSql.Add(sqlHelper.GerarSqlInsertPessoas(indiceLinha, pessoaDict, pessoaID, pessoaFonesDict, consumidorDict, consumidorID, consumidorEnderecoDict, funcionarioDict, enderecoDict));
					}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				salvarArquivo = Tools.GerarNomeArquivo($"PessoasClientes_{estabelecimentoID}_OdontoCompany_Migração", ".sql");
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

		public void ImportarPrecos(string arquivoExcel, int estabelecimentoID, int loginID, string arquivoTabelaPrecosAtuais)
		{
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();
			List<string> linhasSql = [];
			List<string> tabelasAdicionadas = [];

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

			PrecosTabela? precoTabela = null;
			Preco? preco = null;
			ISheet? sheet = null;

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
					long tuss = 0, tabelaID = 0;
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
						
						if (grupoCategoriaDict.ContainsKey(especialidade))
							categoria = (byte)grupoCategoriaDict[especialidade];

						if (!string.IsNullOrEmpty(tabelaIdEcontrada))
							tabelaID = int.Parse(tabelaIdEcontrada);

						else if (!tabelasAdicionadas.Contains(nomeTabela))
						{
							tabelasAdicionadas.Add(nomeTabela);
							precoTabela = new PrecosTabela
							{
								Ativo = ativo,
								DataInclusao = DateTime.Now,
								LoginID = loginID,
								SeguimentoID = 1,
								SolucaoID = 1,
								EstabelecimentoID = estabelecimentoID,
								Nome = nomeTabela
							};
						}

						preco = new Preco
						{
							Ativo = ativo,
							CategoriaID = categoria,
							DataInclusao = DateTime.Now,
							LoginID = loginID,
							Titulo = titulo,
							Valor = valor,
							CodigoTISS = tuss,
							Atalho = abreviacao
						};
					}

					Dictionary<string, object>? precoDict = null;
					Dictionary<string, object>? precoTabelaDict = null;

					if (preco != null)
						precoDict = new Dictionary<string, object>
						{
							{ "LoginID", preco.LoginID },
							{ "Ativo", preco.Ativo },
							{ "CategoriaID", preco.CategoriaID },
							{ "DataInclusao", preco.DataInclusao },
							{ "Titulo", preco.Titulo },
							{ "Valor", preco.Valor },
							{ "CodigoTISS", preco.CodigoTISS },
							{ "Atalho", preco.Atalho }
						};

					if (precoTabela != null)
						precoTabelaDict = new Dictionary<string, object>
						{
							{ "Ativo", precoTabela.Ativo },
							{ "DataInclusao", precoTabela.DataInclusao },
							{ "LoginID", precoTabela.LoginID },
							{ "SeguimentoID", precoTabela.SeguimentoID },
							{ "SolucaoID", precoTabela.SolucaoID },
							{ "EstabelecimentoID", precoTabela.EstabelecimentoID },
							{ "Nome", precoTabela.Nome }
						};

					linhasSql.Add(sqlHelper.GerarSqlInsertPrecos(indiceLinha, precoTabelaDict, tabelaID, precoDict));
				}

				indiceLinha = 0;

				var salvarArquivo = Tools.GerarNomeArquivo($"Precos_{estabelecimentoID}_OdontoCompany_Migração");
				File.WriteAllLines(salvarArquivo + ".sql", linhasSql);

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
			List<string> linhasSql = new();

			//ISheet sheetConsumidores;
			//try
			//{
			//	IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelRecebiveisProd);
			//	sheetConsumidores = workbookConsumidores.GetSheetAt(0);
			//	excelHelper.InitializeDictionaryRecebiveis(sheetConsumidores);
			//}
			//catch (Exception ex)
			//{
			//	throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelRecebiveisProd}\": {ex.Message}");
			//}

			FluxoCaixa fluxoCaixa = null;
			Recebivel recebivel = null;

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", cpf = "";
					string? outroSacadoNome = null, observacaoRecebido = null, observacaoRecebivel = null, documento = null;
					int recibo = 0, codigo = 0;
					int? consumidorID = null, fornecedorID = null, colaboradorID = null, funcionarioID = null, clienteID = null;
					decimal pagoValor = 0, valorOriginal = 0, valorDevido = 0;
					bool isrecebivel = false, isbaixa = false;
					byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
					DateTime dataPagamento = dataHoje, nascimentoData = dataHoje, dataVencimento = dataHoje, dataInclusao = dataHoje, dataBaixa = dataHoje;

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
										observacaoRecebivel = celulaValor;//.ToTipoPagamento()
										break;
									case "ValorOriginal":
										valorOriginal = celulaValor.ArredondarValorV2();
										break;
									case "RecebívelExigível(R/E)":
										isrecebivel = celulaValor == "R" ? true : false;
										break;
									case "DataBaixa":
										dataBaixa = celulaValor.ToData();
										isbaixa = !string.IsNullOrEmpty(celulaValor) ? true : false;
										break;
									case "ValorPago":
										pagoValor = celulaValor.ArredondarValorV2();
										break;
									case "ObservaçãoRecebido":
										observacaoRecebido = celulaValor;
										break;
								}
							}
						}
					}

					if (isrecebivel)
					{
						if (dataVencimento == dataHoje)
							dataVencimento = new DateTime(dataInclusao.Year, dataVencimento.Month, dataVencimento.Day);

						var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: cpf, codigo: codigo.ToString());
						//var consumidorIDValue = excelHelper.GetConsumidorID(cpf: cpf);
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

						if (!string.IsNullOrEmpty(consumidorIDValue))
						{
							if (!excelHelper.RecebivelExists((int)consumidorID, valorOriginal, dataVencimento))
							{
								if (valorOriginal >= 1)
									recebivel = new Recebivel()
									{
										ConsumidorID = consumidorID,
										FornecedorID = fornecedorID,
										ClienteID = clienteID,
										ColaboradorID = colaboradorID,
										SacadoNome = outroSacadoNome,
										EspecieID = (byte)formaPagamento,
										DataEmissao = dataInclusao,
										ValorOriginal = valorOriginal,
										ValorDevido = valorDevido,
										DataBaseCalculo = dataInclusao,
										DataInclusao = dataInclusao,
										DataVencimento = dataVencimento,
										FinanceiroID = respFinanceiroPessoaID,
										LoginID = loginID,
										EstabelecimentoID = estabelecimentoID,
										SituacaoID = (byte)TituloSituacoesID.Normal,
										Observacoes = observacaoRecebivel,
										ExclusaoMotivo = documento,
										DataBaixa = dataBaixa,
										ValorBaixa = pagoValor
										//OrcamentoID
										//PlanoContasID
										//Documento = contratoControle
										//ContratoID = contratoID
									};

								if (pagoValor >= 1)
								{
									var tituloTransacao = TituloTransacoes.Liquidacao;
									if (pagoValor < valorOriginal)
									{
										tituloTransacao = TituloTransacoes.PagamentoParcial;
										valorDevido = valorOriginal - pagoValor;
									}

									fluxoCaixa = new FluxoCaixa()
									{
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
										DevidoValor = valorDevido,
										PagoValor = pagoValor,
										EstabelecimentoID = estabelecimentoID,
										LoginID = loginID,
										DataInclusao = dataBaixa,
										FinanceiroID = respFinanceiroPessoaID,
										Observacoes = observacaoRecebido,
										OutroSacadoNome = outroSacadoNome
									};
								}
							}

							Dictionary<string, object> recebivelDict = null;
							Dictionary<string, object> fluxoCaixaDict = null;

							if (recebivel != null)
								recebivelDict = new Dictionary<string, object>
								{
									{ "ConsumidorID", recebivel.ConsumidorID },
									{ "FornecedorID", recebivel.FornecedorID },
									{ "ClienteID", recebivel.ClienteID },
									{ "ColaboradorID", recebivel.ColaboradorID },
									{ "SacadoNome", recebivel.SacadoNome },
									{ "EspecieID", recebivel.EspecieID },
									{ "DataEmissao", recebivel.DataEmissao },
									{ "ValorOriginal", recebivel.ValorOriginal },
									{ "ValorDevido", recebivel.ValorDevido },
									{ "DataBaseCalculo", recebivel.DataBaseCalculo },
									{ "DataInclusao", recebivel.DataInclusao },
									{ "DataVencimento", recebivel.DataVencimento },
									{ "FinanceiroID", recebivel.FinanceiroID },
									{ "LoginID", recebivel.LoginID },
									{ "EstabelecimentoID", recebivel.EstabelecimentoID },
									{ "SituacaoID", recebivel.SituacaoID },
									{ "Observacoes", recebivel.Observacoes },
									{ "ExclusaoMotivo", recebivel.ExclusaoMotivo },
									{ "ValorBaixa", recebivel.ValorBaixa },
									{ "DataBaixa", recebivel.DataBaixa }
								};

							if (fluxoCaixa != null)
								fluxoCaixaDict = new Dictionary<string, object>
								{
									//{ "ConsumidorID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ConsumidorID).ToArray() },
									//{ "RecebivelID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.RecebivelID).ToArray() },
									//{ "SituacaoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.SituacaoID).ToArray() },
									//{ "PagoMulta", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoMulta).ToArray() },
									//{ "PagoJuros", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoJuros).ToArray() },
									//{ "TipoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TipoID).ToArray() },
									//{ "Data", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Data).ToArray() },
									//{ "TransacaoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TransacaoID).ToArray() },
									//{ "EspecieID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EspecieID).ToArray() },
									//{ "DataBaseCalculo", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataBaseCalculo).ToArray() },
									//{ "DevidoValor", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DevidoValor).ToArray() },
									//{ "PagoValor", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoValor).ToArray() },
									//{ "EstabelecimentoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EstabelecimentoID).ToArray() },
									//{ "LoginID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.LoginID).ToArray() },
									//{ "DataInclusao", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataInclusao).ToArray() },
									//{ "FinanceiroID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.FinanceiroID).ToArray() },
									//{ "Observacoes", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Observacoes).ToArray() },
									//{ "OutroSacadoNome", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.OutroSacadoNome).ToArray() }
									{ "ConsumidorID", fluxoCaixa.ConsumidorID },
									{ "SituacaoID", fluxoCaixa.SituacaoID },
									{ "PagoMulta", fluxoCaixa.PagoMulta },
									{ "PagoJuros", fluxoCaixa.PagoJuros },
									{ "TipoID", fluxoCaixa.TipoID },
									{ "Data", fluxoCaixa.Data },
									{ "TransacaoID", fluxoCaixa.TransacaoID },
									{ "EspecieID", fluxoCaixa.EspecieID },
									{ "DataBaseCalculo", fluxoCaixa.DataBaseCalculo },
									{ "DevidoValor", fluxoCaixa.DevidoValor },
									{ "PagoValor", fluxoCaixa.PagoValor },
									{ "EstabelecimentoID", fluxoCaixa.EstabelecimentoID },
									{ "LoginID", fluxoCaixa.LoginID },
									{ "DataInclusao", fluxoCaixa.DataInclusao },
									{ "FinanceiroID", fluxoCaixa.FinanceiroID },
									{ "Observacoes", fluxoCaixa.Observacoes },
									{ "OutroSacadoNome", fluxoCaixa.OutroSacadoNome }
								};

							if (recebivelDict != null || fluxoCaixaDict != null)
								linhasSql.Add(sqlHelper.GerarSqlInsertRecebiveis(indiceLinha, recebivelDict, fluxoCaixaDict));
						}
					}
				}

				indiceLinha = 0;

				var salvarArquivo = Tools.GerarNomeArquivo($"Recebiveis_{estabelecimentoID}_OdontoCompany_Migração");
				File.WriteAllLines(salvarArquivo + ".sql", linhasSql);

				MessageBox.Show("Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
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