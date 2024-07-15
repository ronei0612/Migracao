using Migracao.DTO;
using Migracao.Models;
using Migracao.Models.DentalOffice;
using Migracao.Models.DTO;
using Migracao.Models.OdontoCompany;
using NPOI.SS.Formula.Functions;
using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace Migracao.Utils
{
    public class ConversorEntidadeParaDTO
    {
        #region Conversores no modelo de importação

        public static List<PacientesDentistasDTO> ConvertPacientesDentistasParaPacientesDentistasDTO(List<Models.Pacientes> pacientes, List<Models.Dentistas>? dentistas = null)
        {
            List<PacientesDentistasDTO> pacientesDentistasDTO = new List<PacientesDentistasDTO>();
            var linha = 1;

            try
            {
                foreach (var paciente in pacientes)
                {
                    var lstPacientes = new PacientesDentistasDTO
                    {
                        Cargo_Clinica = "Paciente",
                        Nome_Completo = paciente.Nome_Paciente.ToNome(),
                        Nome_Social = string.Empty,
                        Apelido = paciente.Nome_Paciente.ToApelido(),
                        CPF = paciente.CPF.ToCPF(),
                        Observacoes = paciente.Observacoes,
                        Email = paciente.E_mail.ToEmail(),
                        RG = paciente.RG.ToUpper(),
                        Sexo = paciente.Sexo.ToSexo("m", "f") ? "Masculino" : "Feminino",
                        Data_Nascimento = paciente.Data_de_Nascimento.ToDataNull().ToString(),
                        Telefone_Principal = paciente.Telefone_Principal.ToFone().ToString(),
                        Celular = paciente.Celular.ToFone().ToString(),
                        Telefone_Alternativo = paciente.Telefone_Alternativo.ToFone().ToString(),
                        Logradouro = paciente.Logradouro,
                        Numero = paciente.Numero,
                        Bairro = paciente.Bairro,
                        Cidade = paciente.Cidade.ToCidade(paciente.UF),
                        UF = paciente.UF?.ToUpper() ?? "",
                        CEP = paciente.CEP,
                        Codigo_Conselho_Estado = string.Empty,
                        Estado_Civil = string.Empty,
                        Cidade_Nascimento = string.Empty,
                        Profissao = string.Empty
                    };

                    pacientesDentistasDTO.Add(lstPacientes);
                    linha++;
                };

                //foreach (var dentista in dentistas)
                //{
                //    var lstPacientes = new PacientesDentistasDTO
                //    {
                //        Cargo_Clinica = "Dentista",
                //        Nome_Completo = dentista.Nome_Completo.ToNome() ?? dentista.Apelido.ToApelido(),
                //        Nome_Social = string.Empty,
                //        Apelido = dentista.Apelido.ToDentista(dentista.Nome_Completo),
                //        Observacoes = dentista.Observacoes,
                //        Email = dentista.Email,
                //        Telefone_Principal = dentista.Telefone.ToFone().ToString(),
                //        Codigo_Conselho_Estado = dentista.Codigo_do_Conselho_e_Estado
                //    };

                //    pacientesDentistasDTO.Add(lstPacientes);
                //    linha++;
                //};
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Pacientes (Linha {linha}): {error.Message}");
            }

            return pacientesDentistasDTO;
        }

        public static List<AgendamentosDTO> ConvertAgendamentodsParaAgendamentosDTO(List<Agendamentos> agendamentos)
        {
            List<AgendamentosDTO> lstAgendamentosDTO = new List<AgendamentosDTO>();

            try
            {
                foreach (var agendamento in agendamentos)
                {
                    var minutos = agendamento.Hora.Split(':')[1];
                    var horas = agendamento.Hora.Split(':')[0];
                    var dataInicio = agendamento.Data;

                    if (!string.IsNullOrEmpty(horas))
                        dataInicio = dataInicio.AddHours(double.Parse(horas));
                    if (!string.IsNullOrEmpty(minutos))
                        dataInicio = dataInicio.AddMinutes(double.Parse(minutos));

                    var dataTermino = dataInicio;
                    //var idsEncontrados = agendamentos.Where(agenda => agenda.Equals(agendamento.ID)).Count();
                    //if (idsEncontrados > 0)
                    //    dataTermino = dataTermino.AddMinutes(15 * idsEncontrados);
                    //else
                    //    dataTermino = dataTermino.AddMinutes(15);

                    var lstAgendamento = new AgendamentosDTO
                    {
                        Lancamento = agendamento.Lancamento,
                        CPF = agendamento.Paciente_CPF.ToCPF(),
                        Nome_Completo = agendamento.Nome.ToNome(),
                        Telefone = agendamento.Telefone,
                        Data_Inicio = dataInicio,
                        Data_Termino = dataTermino,
                        Data_Inclusao = agendamento.Data_Inclusao,
                        Nome_Completo_Dentista = agendamento.Nome_Dentista,
                        Observacao = agendamento.Observacao
                    };

                    lstAgendamentosDTO.Add(lstAgendamento);
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Agendamentos: {error.Message}");
            }

            return lstAgendamentosDTO;
        }

        public static List<DesenvolvimentoClinicoDTO> ConvertDesenvolvimentoClinicoParaDesenvolvimentoClinicoDTO(List<DesenvolvimentoClinico> desenvClicnicos, List<Agendamentos> agendamentos)
        {
            List<DesenvolvimentoClinicoDTO> lstDesenvolvimentoClinicoDTO = new List<DesenvolvimentoClinicoDTO>();
            var linha = 1;

            try
            {

                Parallel.ForEach(agendamentos, agendamento =>
                {
                    var minutos = agendamento.Hora.Split(':')[1];
                    var horas = agendamento.Hora.Split(':')[0];
                    var dataInicio = agendamento.Data;

                    if (!string.IsNullOrEmpty(horas))
                        dataInicio = dataInicio.AddHours(double.Parse(horas));

                    if (!string.IsNullOrEmpty(minutos))
                        dataInicio = dataInicio.AddMinutes(double.Parse(minutos));

                    var dataTermino = dataInicio;

                    var numLancamentos = agendamentos
                                        .Where(m => m.Lancamento == agendamento.Lancamento)
                                        .Count();

                    var minutosParaAdicionar = numLancamentos > 0 ? 15 * numLancamentos : 15;
                    dataTermino = dataTermino.AddMinutes(minutosParaAdicionar);

                    var desenvolvimentoClinico = new DesenvolvimentoClinicoDTO
                    {
                        CPF = agendamento.Paciente_CPF.ToCPF(),
                        Nome_Completo = agendamento.Nome.ToNome(),
                        Telefone = agendamento.Telefone.ToFone().ToString(),
                        Dentista = agendamento.Nome_Dentista.ToNome(),
                        //ID = agendamento.Codigo_Responsavel,
                        Desenvolvimento_Clinico = agendamento.Observacao,
                        Data_Hora_Inicio = agendamento.Data_Inclusao.ToString(),
                        Data_Hora_Termino = dataTermino.ToString(),
                        Data_Hora_Atendimento_Inicio = agendamento.Data.ToString(),
                        Data_Hora_Atendimento_Termino = dataTermino.ToString(),
                        //Observacao = agendamento.Observacao
                    };

                    lock (lstDesenvolvimentoClinicoDTO)
                    {
                        lstDesenvolvimentoClinicoDTO.Add(desenvolvimentoClinico);
                    }
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Desenvolvimento Clinico (Linha {linha}): {error.Message}");
            }

            return lstDesenvolvimentoClinicoDTO;
        }

        public static List<ProcedimentosManutencaoDTO> ConvertProcedimentosParaProcedimentosDTO(List<Models.Procedimentos> procedimentos, List<Manutencoes> manutencoes)
        {
            List<ProcedimentosManutencaoDTO> lstProcedManutDTO = new List<ProcedimentosManutencaoDTO>();
            decimal? valorTotal = 0;
            int docsEncontrados = 0;
            var linha = 1;

            try
            {

                foreach (var procedimento in procedimentos)
                {
                    var lstProcedManut = new ProcedimentosManutencaoDTO
                    {
                        Numero_Controle = procedimento.Numero_Controle,
                        Paciente_CPF = procedimento.Paciente_CPF.ToCPF(),
                        Paciente_Nome = procedimento.Nome_Paciente.ToNome(),
                        Dentista_CPF = procedimento.Dentista_CPF.ToCPF(),
                        Dentista_Nome = procedimento.Dentista_Nome.ToNome(),
                        Dente = procedimento.Dente,
                        Procedimento_Nome = procedimento.NOME_PRODUTO,
                        Procedimento_Valor = procedimento.Valor.ToMoeda().ToString(),
                        Procedimento_Observacao = procedimento.Observacao,
                        Data_Inicio = procedimento.Data_Inicio.ToDataNull().ToString(),
                        Data_Termino = procedimento.Data_Termino.ToDataNull().ToString(),
                        Data_Atendimento = procedimento.Data_Atendimento.ToDataNull().ToString(),
                        //Valor_Original = procedimento.Valor_Original.ToString(),
                        //Valor_Pagamento = procedimento.Valor_Pagamento.ToString(),
                        //Data_Pagamento = procedimento.Data_Pagamento.ToString(),
                    };

                    lstProcedManutDTO.Add(lstProcedManut);
                    linha++;
                };

                linha = 1;

                foreach (var manutencao in manutencoes)
                {
                    var contaQtdOrto = manutencoes
                                     .Where(m => m.Numero_Controle == manutencao.Numero_Controle)
                                     .Count();

                    var listaValores = manutencoes.Where(linha => linha.Paciente_CPF.Equals(manutencao.Paciente_CPF)).ToList();

                    valorTotal = ConverterHelper.SomarValores(listaValores, m => m.Valor_Devido);

                    var lstManutencao = new ProcedimentosManutencaoDTO
                    {
                        Numero_Controle = manutencao.Numero_Controle,
                        Paciente_CPF = manutencao.Paciente_CPF.ToCPF(),
                        Paciente_Nome = manutencao.Nome_Paciente.ToNome(),
                        Dentista_Nome = manutencao.Dentista_Nome.ToNome(),
                        Procedimento_Nome = manutencao.Procedimento_Nome,
                        Procedimento_Valor = manutencao.Procedimento_Valor.ToMoeda().ToString(),
                        Valor_Original = manutencao.Valor_Original.ToString(),
                        Valor_Pago = manutencao.Valor_Pagamento.ToString(),
                        Data_Pagamento = manutencao.Data_Pagamento.ToString(),
                        Procedimento_Observacao = manutencao.Procedimentos_Observacao,
                        Quantidade_Orto = contaQtdOrto.ToString(),
                        Tipo_Pagamento = manutencao.Tipo_Pagamento,
                        Vencimento = manutencao.Vencimento.ToString(),
                        Valor_Devido = manutencao.Valor_Devido?.ToString(),
                        Valor_Total = valorTotal.ToString(),
                        Data_Atendimento = manutencao.Data_Atendimento.ToString()
                    };

                    lstProcedManutDTO.Add(lstManutencao);

                    linha++;
                    valorTotal = 0;
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos e Manutenções (Linha {linha}): {error.Message}");
            }

            return lstProcedManutDTO;
        }

        public static List<ProcedimentosPrecosDTO> ConvertProcedimentosPrecosParaProcedimentosPrecosDTO(List<ProcedimentosPrecos> gruposProcedimentos)
        {
            List<ProcedimentosPrecosDTO> lstGruposProcedimentosDTO = new List<ProcedimentosPrecosDTO>();

            try
            {
                foreach (var grupoProcedimento in gruposProcedimentos)
                {
                    var lstGruposProcedimentos = new ProcedimentosPrecosDTO
                    {
                        Nome = grupoProcedimento.Procedimento_Nome,
                        Tabela = grupoProcedimento.Tabela,
                        Especialidade = grupoProcedimento.Nome_Grupo,
                        NomeProcedimento = grupoProcedimento.Procedimento_Nome,
                        Abreviacao = grupoProcedimento.Abreviacao,
                        Preco = grupoProcedimento.Preco.ToString(),
                        TUSS = grupoProcedimento.TUSS
                    };

                    lstGruposProcedimentosDTO.Add(lstGruposProcedimentos);
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos Tabela Preços: {error.Message}");
            }

            return lstGruposProcedimentosDTO;
        }

        public static List<ManutencoesDTO> ConvertManutencoesParaManutencoesDTO(List<Manutencoes> manutencoes)
        {
            List<ManutencoesDTO> lstManutencoesDTO = new List<ManutencoesDTO>();
            decimal? valorTotal = 0;
            int docsEncontrados = 0;

            try
            {
                foreach (var manutencao in manutencoes)
                {
                    var somaValores = manutencoes
                                     .Where(m => m.Numero_Controle == manutencao.Numero_Controle)
                                     .Count();


                    var listaValores = manutencoes.Where(linha => linha.Paciente_CPF.Equals(manutencao.Paciente_CPF)).ToList();

                    var selecionaLinha = listaValores.Select(linha => linha.Valor_Devido?.Replace(",", "."));

                    foreach (var item in selecionaLinha)
                    {
                        if (!string.IsNullOrEmpty(item))
                            valorTotal += Convert.ToDecimal(item, CultureInfo.InvariantCulture);
                    }

                    var lstManutencao = new ManutencoesDTO
                    {
                        Numero_Controle = manutencao.Numero_Controle,
                        Paciente_CPF = manutencao.Paciente_CPF,
                        Paciente_Nome = manutencao.Nome_Paciente,
                        Dentista_Nome = manutencao.Dentista_Nome,
                        Procedimento_Nome = manutencao.Procedimento_Nome,
                        Procedimento_Valor = manutencao.Procedimento_Valor,
                        Valor_Original = manutencao.Valor_Original.ToString(),
                        Valor_Pagamento = manutencao.Valor_Pagamento.ToString(),
                        Data_Pagamento = manutencao.Data_Pagamento.ToString(),
                        Dente = manutencao.Dente,
                        Procedimento_Observacao = manutencao.Procedimentos_Observacao,
                        Quantidade_Orto = somaValores.ToString(),
                        Tipo_Pagamento = manutencao.Tipo_Pagamento,
                        Vencimento = manutencao.Vencimento.ToString(),
                        Valor_Devido = manutencao.Valor_Devido?.ToString(),
                        Valor_Total = valorTotal.ToString(),
                        Data_Atendimento = manutencao.Data_Atendimento.ToString()
                    };

                    lstManutencoesDTO.Add(lstManutencao);

                    valorTotal = 0;
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Manutencao: {error.Message}");
            }

            return lstManutencoesDTO;
        }


        #endregion



        #region Financeiro

        public static List<FinanceiroRecebiveisDTO> ConvertRecebiveisParaRecebiveisDTO(List<Recebivel> recebiveis)
        {
            List<FinanceiroRecebiveisDTO> lstRecebiveisDTO = new List<FinanceiroRecebiveisDTO>();

            try
            {
                foreach (var recebivel in recebiveis)
                {
                    var lstRecebiveis = new FinanceiroRecebiveisDTO
                    {
                        Paciente_CPF = recebivel.CNPJ_CPF,
                        Nome = recebivel.Nome,
                        Numero_Controle = recebivel.Numero_Controle,
                        Observacao_Recebivel = recebivel.Observacao,
                        Recebivel_Exigivel = "R",
                        Valor_Devido = recebivel.Valor_Devido.ToString().ArredondarValorV2().ToString(),
                        Data_Vencimento = recebivel.Data_Vencimento.ToString("dd/MM/yyyy"),
                        Emissao = recebivel.Emissao.ToString("dd/MM/yyyy"),
                        Duplicata = recebivel.Duplicata.ToString(),
                        Parcela = recebivel.Parcela.ToString(),
                        Tipo_Pagamento = recebivel.Tipo_Pagamento.ToString(),
                        Valor_Original = recebivel.Valor_Original.ToString().ArredondarValorV2().ToString(),
                        Vencimento_Recebivel = recebivel.Vencimento_Original.ToString("dd/MM/yyyy"),
                        Situacao = recebivel.Situacao,
                        Nome_Grupo = recebivel.Nome_Grupo,
                        Ordem = recebivel.Ordem.ToString()
                    };

                    lstRecebiveisDTO.Add(lstRecebiveis);
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos: {error.Message}");
            }

            return lstRecebiveisDTO;
        }

        public static List<PagosExigiveisDTO> ConvertRecebidosParaRecebidosDTO(List<Recebidos> recebidos)
        {
            List<PagosExigiveisDTO> lstReceberDTO = new List<PagosExigiveisDTO>();

            try
            {
                foreach (var receber in recebidos)
                {

                    var tipoPagamento = receber.Tipo_Pagamento;
                    string formaPagamento = ExcelHelper.GetEspecieIDFromFormaPagamentoEntidades(tipoPagamento);

                    var lstReceber = new PagosExigiveisDTO
                    {
                        CPF = receber.CNPJ_CPF.ToCPF(),
                        Nome = receber.Nome_Paciente.ToCPF(),
                        Numero_Controle = receber.Numero_Controle,
                        Recebivel_Exigivel = "R",
                        Valor_Devido = receber.Valor_Devido.ToString(),
                        Valor_Pago = receber.Valor_Pago.ToString(),
                        Data_Vencimento = receber.Data_Vencimento.ToShortDateString(),
                        Data_Pagamento = receber.Data_Baixa.ToShortDateString(),
                        Observacao_Recebido = ("Observação: " + receber.Observacao + " | Documento: " + receber.Tipo_Documento + " | Situação: " + receber.Situacao +
                        " | Nome do Grupo: " + receber.Nome_Grupo + " | Ordem: " + receber.Ordem),
                        Tipo_Pagamento = formaPagamento,
                        Valor_Original = receber.Valor_Original.ToString(),
                        //Vencimento_Recebivel = receber.Vencimento_Recebivel.ToShortDateString(),
                        //Duplicata = receber.Duplicata,
                        Parcela = receber.Parcela.ToString(),
                        //Tipo_Especie_Pagamento = receber.Tipo_Especie,
                        //Especie_Pagamento = ExcelHelper.GetEspecieIDFromFormaPagamento(receber.Tipo_Especie).ToString(),
                        Pagamento_Observacoes = receber.Pagamento_Observacoes
                    };

                    lstReceberDTO.Add(lstReceber);
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos: {error.Message}");
            }

            return lstReceberDTO;
        }

        #endregion

        #region Complementares

        public static List<ProcedimentosDTO> ConvertProcedimentosParaProcedimentosDTO(List<Models.Procedimentos> procedimentos)
        {
            List<ProcedimentosDTO> lstProcedimentosDTO = new List<ProcedimentosDTO>();

            try
            {

                foreach (var procedimento in procedimentos)
                {
                    var lstProcedimento = new ProcedimentosDTO
                    {
                        Numero_Controle = procedimento.Numero_Controle,
                        Paciente_CPF = procedimento.Paciente_CPF.ToCPF(),
                        Paciente_Nome = procedimento.Nome_Paciente,
                        Dentista_CPF = procedimento.Dentista_CPF,
                        Dentista_Nome = procedimento.Dentista_Nome,
                        Dente = procedimento.Dente,
                        Procedimento_Nome = procedimento.NOME_PRODUTO,
                        Procedimento_Valor = procedimento.Valor,
                        Procedimento_Observacao = procedimento.Observacao,
                        Data_Inicio = procedimento.Data_Inicio.ToData().ToShortDateString(),
                        Data_Termino = procedimento.Data_Termino,
                        Data_Atendimento = procedimento.Data_Atendimento
                    };

                    lstProcedimentosDTO.Add(lstProcedimento);
                }
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos: {error.Message}");
            }

            return lstProcedimentosDTO;
        }
        public static List<DentistasDTO> ConvertDentistasParaDentistasDTO(List<Dentistas> dentistas)
        {
            List<DentistasDTO> lstDentistasDTO = new List<DentistasDTO>();

            try
            {
                foreach (var dentista in dentistas)
                {
                    var lstDentistas = new DentistasDTO
                    {
                        //Cargo_Clinica = "Dentista",
                        //Nome = dentista.Nome_Completo,
                        //Nome_Social = string.Empty,
                        //Nome_Completo = dentista.Nome_Completo,
                        //Apelido = dentista.Nome_Completo.GetPrimeirosCaracteres(20),
                        //Observacoes = dentista.Observacoes,
                        //Email = dentista.Email,
                        //Telefone_Principal = dentista.Telefone,
                        //Codigo_Conselho_Estado = dentista.Codigo_do_Conselho_e_Estado
                    };

                    lstDentistasDTO.Add(lstDentistas);
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
            }

            return lstDentistasDTO;
        }
        public static List<RecebiveisHistVendaDTO> ConvertRecebiveisHistVendaParaRecebiveisHistVendaDTO(List<RecebiveisHistVenda> recHistVendas)
        {
            List<RecebiveisHistVendaDTO> lstRecebiveisHistVendaDTO = new List<RecebiveisHistVendaDTO>();

            try
            {
                foreach (var recHistVenda in recHistVendas)
                {
                    var lstRecHistVendas = new RecebiveisHistVendaDTO
                    {
                        CPF = recHistVenda.Paciente_CPF,
                        Nome = recHistVenda.Nome_Paciente,
                        Observacao_Recebivel = recHistVenda.Observacao_Recebivel,
                        Documento_Ref = recHistVenda.Documento_Ref,
                        Valor_Original = recHistVenda.Valor_Original.ToString(),
                        Vencimento = recHistVenda.Vencimento.ToString(),
                        Emissao = recHistVenda.Emissao,
                        Recebivel_Exigivel = recHistVenda.Recebivel_Exigivel
                    };

                    lstRecebiveisHistVendaDTO.Add(lstRecHistVendas);
                };
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Recebiveis Histórico Venda: {error.Message}");
            }

            return lstRecebiveisHistVendaDTO;
        }

        #endregion






    }
}
