using Migracao.Models;
using Migracao.Utils;
using NPOI.SS.UserModel;
using System.Runtime.ConstrainedExecution;

namespace Migracao.Sistems
{
    internal class OdontoCompany
    {
        public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, int estabelecimentoID, int respFinanceiroPessoaID, int loginID)
        {
            var dataHoje = DateTime.Now;
            var indiceLinha = 0;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            ISheet sheetConsumidores;
            try
            {
                IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelConsumidores);
                sheetConsumidores = workbookConsumidores.GetSheetAt(0);
                excelHelper.InitializeDictionary(sheetConsumidores);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
            }

            var fluxoCaixas = new List<FluxoCaixa>();

            try
            {
                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

                    string nomeCompleto = "", outroSacadoNome = "";
                    int controle = 0, recibo = 0, codigo = 0;
                    int? consumidorID = 0;
                    decimal pagoValor = 0;
                    byte titulosEspecies = 0;
                    var dataPagamento = dataHoje;
                    var nascimentoData = dataHoje;

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
                                    case "paciente":
                                        nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                        break;
                                    case "numero_registro":
                                        codigo = int.Parse(celulaValor);
                                        break;
                                    case "data_pagamento":
                                        dataPagamento = celulaValor.ToData();
                                        break;
                                    case "forma_pagamento":
                                        celulaValor.ToTipoPagamento();
                                        break;
                                    case "valor":
                                        pagoValor = celulaValor.ToMoeda();
                                        break;
                                }
                            }
                        }
                    }

                    var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, codigo: codigo.ToString());
                    if (!string.IsNullOrEmpty(consumidorIDValue))
                    {
                        consumidorID = int.Parse(consumidorIDValue);
                        outroSacadoNome = "null";

                        fluxoCaixas.Add(new FluxoCaixa()
                        {
                            ConsumidorID = consumidorID,
                            SituacaoID = 1,
                            PagoMulta = 0,
                            PagoJuros = 0,
                            TipoID = (byte)TransacaoTiposID.Recebimento,
                            Data = dataPagamento,
                            TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
                            EspecieID = titulosEspecies,
                            DataBaseCalculo = dataPagamento,
                            DevidoValor = pagoValor,
                            PagoValor = pagoValor,
                            EstabelecimentoID = estabelecimentoID,
                            LoginID = loginID,
                            DataInclusao = dataPagamento,
                            FinanceiroID = respFinanceiroPessoaID
                        });
                    }
                    else
                    {
                        consumidorID = null;
                        outroSacadoNome = nomeCompleto.GetPrimeirosCaracteres(50);

                        fluxoCaixas.Add(new FluxoCaixa()
                        {
                            OutroSacadoNome = nomeCompleto.GetPrimeirosCaracteres(50),
                            SituacaoID = 1,
                            PagoMulta = 0,
                            PagoJuros = 0,
                            TipoID = (byte)TransacaoTiposID.Recebimento,
                            Data = dataPagamento,
                            TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
                            EspecieID = titulosEspecies,
                            DataBaseCalculo = dataPagamento,
                            DevidoValor = pagoValor,
                            PagoValor = pagoValor,
                            EstabelecimentoID = estabelecimentoID,
                            LoginID = loginID,
                            DataInclusao = dataPagamento,
                            FinanceiroID = respFinanceiroPessoaID
                        });
                    }
                }

                indiceLinha = 0;

                var dados = new Dictionary<string, object[]>
                {
                    { "ConsumidorID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ConsumidorID).ToArray() },
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
                    { "OutroSacadoNome", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.OutroSacadoNome).ToArray() }
                };

				var salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_FluxoCaixa");
				sqlHelper.GerarSqlInsert("_MigracaoFluxoCaixa_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
				Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
            }
            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

        public void ImportarFornecedores(string arquivoExcel, string arquivoExcelCidades, int estabelecimentoID, int loginID)
        {
            var dataHoje = DateTime.Now;
            var indiceLinha = 0;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

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
                //var consumidores = new List<Consumidor>();
                var fornecedores = new List<Fornecedor>();

                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;
                    bool cliente = false, fornecedor = false;
                    DateTime dataNascimento, dataCadastro = dataHoje;
                    int numcadastro;
                    string nomeCompleto = "", cpf = "", rg = "", email = "", apelido = "";
                    bool sexo = true;
                    long telefonePrinc, telefoneAltern, telefoneComercial, telefoneOutro, celular;

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
                                        nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                        apelido = celulaValor.GetPrimeiroNome();
                                        break;
                                    case "CGC_CPF":
                                        cpf = celulaValor.ToCPF();
                                        break;
                                    case "INSC_RG":
                                        rg = celulaValor.GetPrimeirosCaracteres(20);
                                        break;
                                    case "SEXO_M_F":
                                        var sexoLetra = celulaValor.ToLower();
                                        sexo = sexoLetra == "m" || sexoLetra != "f";
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
                                    //case "ENDERECO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "BAIRRO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "NUM_ENDERECO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "CIDADE":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "ESTADO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "CEP":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "OBS1":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "NUM_CONVENIO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    case "DT_CADASTRO":
                                        dataCadastro = celulaValor.ToData();
                                        break;
                                    case "DT_NASCIMENTO":
                                        dataNascimento = celulaValor.ToData();
                                        break;
                                }
                            }
                        }
                    }

                    if (fornecedor)
                        fornecedores.Add(new Fornecedor()
                        {
                            Ativo = true,
                            DataInclusao = dataCadastro,
                            EstabelecimentoID = estabelecimentoID,
                            NomeFantasia = nomeCompleto,
                            LoginID = loginID
                        });
                }

                indiceLinha = 0;

                var dados = new Dictionary<string, object[]>
                {
                    { "Ativo", fornecedores.ConvertAll(fornecedor => (object)fornecedor.Ativo).ToArray() },
                    { "DataInclusao", fornecedores.ConvertAll(fornecedor => (object)fornecedor.DataInclusao).ToArray() },
                    { "EstabelecimentoID", fornecedores.ConvertAll(fornecedor => (object)fornecedor.EstabelecimentoID).ToArray() },
                    { "LoginID", fornecedores.ConvertAll(fornecedor => (object)fornecedor.LoginID).ToArray() },
                    { "NomeFantasia", fornecedores.ConvertAll(fornecedor => (object)fornecedor.NomeFantasia).ToArray() }
                };

				var salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_Consumidores");
				sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
				Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
            }

            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

		public void ImportarPacientes(string arquivoExcel, string arquivoExcelCidades, int estabelecimentoID, int loginID)
        {
            var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            DateTime dataHoje = DateTime.Now;
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

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
                var consumidores = new List<Consumidor>();
                var pessoas = new List<Pessoa>();
                var consumidoresEnderecos = new List<ConsumidorEndereco>();
                var pessoaFones = new List<PessoaFone>();
				var fornecedores = new List<Fornecedor>();
				var empresas = new List<Empresa>();
                var empresasEnderecos = new List<Endereco>();
				var fornecedorFones = new List<FornecedorFone>();

				foreach (var linha in excelHelper.linhas)
                {
                    bool cliente = false, fornecedor = false;
                    DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
                    int pessoaID = 0, cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
                    long telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string nomeCompleto = "null", documento = "null", rg = "null", email = "null", apelido = "null", nascimentoLocal = "null", profissaoOutra = "null", logradouro = "null",
						 complemento = "null", bairro = "null", logradouroNum = "null", numcadastro = "null", cidade = "null", estado = "null", observacao = "null";

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
                                        nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                        apelido = celulaValor.GetPrimeiroNome();
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
                                        logradouro = celulaValor;
                                        break;
                                    case "BAIRRO":
                                        bairro = celulaValor;
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

                    if (documento.IsCPF())
                    {
                        pessoas.Add(new Pessoa()
                        {
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
                            ConselhoCodigo = indiceLinha.ToString()
                        });

                        if (fornecedor)
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

							if (!string.IsNullOrWhiteSpace(cidade))
							{
								var cidadeID = excelHelper.GetCidadeID(cidade, estado);
								empresasEnderecos.Add(new Endereco()
								{
									Ativo = true,
									Cep = cep,
									CidadeID = cidadeID,
									DataInclusao = dataCadastro,
									EnderecoTipoID = (byte)EnderecoTipos.Residencial,
									Logradouro = logradouro,
									LogradouroNum = logradouroNum,
									Bairro = bairro,
									LogradouroTipoID = (int)LogradouroTipos.Outros,
									Complemento = complemento,
									ParentID = indiceLinha,
									TableID = 1,
								});
							}

							if (celular > 0)
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Celular,
									Telefone = celular,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefonePrinc > 0)
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Principal,
									Telefone = telefonePrinc,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefoneAltern > 0)
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Alternativo,
									Telefone = telefoneAltern,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefoneComercial > 0)
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Comercial,
									Telefone = telefoneComercial,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});

							if (telefoneOutro > 0)
								fornecedorFones.Add(new FornecedorFone()
								{
									FornecedorID = indiceLinha,
									FoneTipoID = (short)FoneTipos.Outros,
									Telefone = telefoneOutro,
									DataInclusao = dataCadastro,
									LoginID = loginID
								});
						}

                        else if (!fornecedor)
						{
							consumidores.Add(new Consumidor()
							{
								Ativo = true,
								DataInclusao = dataCadastro,
								EstabelecimentoID = estabelecimentoID,
								LGPDSituacaoID = 0,
								LoginID = loginID,
								PessoaID = indiceLinha,
								CodigoAntigo = numcadastro,
								Observacoes = observacao
							});

							if (!string.IsNullOrWhiteSpace(cidade))
							{
								var cidadeID = excelHelper.GetCidadeID(cidade, estado);
								consumidoresEnderecos.Add(new ConsumidorEndereco()
								{
									Ativo = true,
									ConsumidorID = indiceLinha,
									EnderecoTipoID = (short)EnderecoTipos.Residencial,
									LogradouroTipoID = (int)LogradouroTipos.Outros,
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

							if (celular > 0)
                                pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Celular,
                                    Telefone = celular,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefonePrinc > 0)
                                pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Principal,
                                    Telefone = telefonePrinc,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneAltern > 0)
                                pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Alternativo,
                                    Telefone = telefoneAltern,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneComercial > 0)
                                pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Comercial,
                                    Telefone = telefoneComercial,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneOutro > 0)
                                pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Outros,
                                    Telefone = telefoneOutro,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });
						}
					}
                    else if (documento.IsCNPJ_CGC())
                    {
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
					{ "ConselhoCodigo", pessoas.ConvertAll(pessoa => (object)pessoa.ConselhoCodigo).ToArray() }
				};

                salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_Pessoas");
                sqlHelper.GerarSqlInsert("_MigracaoPessoas_Temp", salvarArquivo, pessoasDict);
                excelHelper.GravarExcel(salvarArquivo, pessoasDict);
				Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");


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

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_Consumidores");
				sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", salvarArquivo, consumidoresDict);
				excelHelper.GravarExcel(salvarArquivo, consumidoresDict);


				var consumidoresEnderecosDict = new Dictionary<string, object[]>
				{
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

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_ConsumidorEnderecos");
				sqlHelper.GerarSqlInsert("_MigracaoConsumidorEnderecos_Temp", salvarArquivo, consumidoresEnderecosDict);
				excelHelper.GravarExcel(salvarArquivo, consumidoresEnderecosDict);
				

				var pessoaFonesDict = new Dictionary<string, object[]>
				{
					{ "PessoaID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.PessoaID).ToArray() },
					{ "FoneTipoID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.FoneTipoID).ToArray() },
					{ "Telefone", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.Telefone).ToArray() },
					{ "DataInclusao", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.DataInclusao).ToArray() },
					{ "LoginID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_PessoaFones");
				sqlHelper.GerarSqlInsert("_MigracaoPessoaFones_Temp", salvarArquivo, pessoaFonesDict);
				excelHelper.GravarExcel(salvarArquivo, pessoaFonesDict);


				var fornecedorFonesDict = new Dictionary<string, object[]>
				{
					{ "FornecedorID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.FornecedorID).ToArray() },
					{ "FoneTipoID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.FoneTipoID).ToArray() },
					{ "Telefone", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.Telefone).ToArray() },
					{ "DataInclusao", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.DataInclusao).ToArray() },
					{ "LoginID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_FornecedorFones");
				sqlHelper.GerarSqlInsert("_MigracaoFornecedorFones_Temp", salvarArquivo, fornecedorFonesDict);
				excelHelper.GravarExcel(salvarArquivo, fornecedorFonesDict);


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

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_Empresas");
				sqlHelper.GerarSqlInsert("_MigracaoEmpresas_Temp", salvarArquivo, empresasDict);
				excelHelper.GravarExcel(salvarArquivo, empresasDict);


				var empresasEnderecosDict = new Dictionary<string, object[]>
				{
	                { "Ativo", empresasEnderecos.ConvertAll(endereco => (object)endereco.Ativo).ToArray() },
	                { "Cep", empresasEnderecos.ConvertAll(endereco => (object)endereco.Cep).ToArray() },
	                { "CidadeID", empresasEnderecos.ConvertAll(endereco => (object)endereco.CidadeID).ToArray() },
	                { "DataInclusao", empresasEnderecos.ConvertAll(endereco => (object)endereco.DataInclusao).ToArray() },
	                { "EnderecoTipoID", empresasEnderecos.ConvertAll(endereco => (object)endereco.EnderecoTipoID).ToArray() },
	                { "Logradouro", empresasEnderecos.ConvertAll(endereco => (object)endereco.Logradouro).ToArray() },
	                { "LogradouroNum", empresasEnderecos.ConvertAll(endereco => (object)endereco.LogradouroNum).ToArray() },
	                { "Bairro", empresasEnderecos.ConvertAll(endereco => (object)endereco.Bairro).ToArray() },
	                { "LogradouroTipoID", empresasEnderecos.ConvertAll(endereco => (object)endereco.LogradouroTipoID).ToArray() },
	                { "Complemento", empresasEnderecos.ConvertAll(endereco => (object)endereco.Complemento).ToArray() },
	                { "ParentID", empresasEnderecos.ConvertAll(endereco => (object)endereco.ParentID).ToArray() },
	                { "TableID", empresasEnderecos.ConvertAll(endereco => (object)endereco.TableID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_Enderecos");
				sqlHelper.GerarSqlInsert("_MigracaoEnderecos_Temp", salvarArquivo, empresasEnderecosDict);
				excelHelper.GravarExcel(salvarArquivo, empresasEnderecosDict);
			}

            catch (Exception error)
            {
				throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }
    }
}
