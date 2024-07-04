using Migracao.DTO;
using Migracao.Models;
using Migracao.Models.DTO;
using NPOI.SS.Formula.Functions;
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
        public static List<AgendamentosDTO> ConvertAgendamentosDTOParaAgendamentosDTO(List<Agendamentos> agendamentos)
        {
            List<AgendamentosDTO> agendamentosDTO = new List<AgendamentosDTO>();

            try
            {
                Parallel.ForEach(agendamentos, agendamento =>
                {
                    var lstAgendamentos = new AgendamentosDTO
                    {
                       
                    };

                    agendamentosDTO.Add(lstAgendamentos);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Agendamentos: {error.Message}");
            }

            return agendamentosDTO;
        }

        public static List<PacientesDTO> ConvertPacientesParaPacientesDTO(List<Pacientes> pacientes)
        {
            List<PacientesDTO> pacientesDTO = new List<PacientesDTO>();

            try
            {
                Parallel.ForEach(pacientes, paciente =>
                {
                    var lstPacientes = new PacientesDTO
                    {
                        Codigo = string.Empty,
                        Ativo = "R",
                        Nome_Completo = paciente.NOME,
                        Nome_Social = string.Empty,
                        Apelido = paciente.NOME.GetPrimeirosCaracteres(20).ToNome(),
                        Documento = paciente.CGC_CPF.ToCPF(),
                        Data_Cadastro = paciente.DT_CADASTRO,
                        Observacoes = paciente.OBS1,
                        Email = paciente.EMAIL,
                        RG = paciente.INSC_RG.GetPrimeirosCaracteres(20),
                        Sexo = paciente.SEXO_M_F.ToSexo("m", "f") ? "M" : "F",
                        Nascimento_Data = paciente.DT_NASCIMENTO,
                        Paciente = paciente.CLIENTE,
                        Funcionario = "N",
                        Fornecedor = paciente.FORNECEDOR,
                        Telefone_Principal = paciente.FONE1,
                        Celular = paciente.CELULAR,
                        Telefone_Alternativo = paciente.FONE2,
                        Logradouro = paciente.ENDERECO.PrimeiraLetraMaiuscula(),
                        Logradouro_Num = paciente.NUM_ENDERECO,
                        Bairro = paciente.BAIRRO,
                        Cidade = paciente.CIDADE,
                        Estado = paciente.ESTADO,
                        CEP = paciente.CEP,
                    };

                    pacientesDTO.Add(lstPacientes);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
            }

            return pacientesDTO;
        }

        public static List<ManutencoesDTO> ConvertManutencoesParaManutencoesDTO(List<Manutencoes> manutencoes)
        {
            List<ManutencoesDTO> lstManutencoesDTO = new List<ManutencoesDTO>();

            //var listaValores = linhas.Where(linha => linha[1].Equals(cpf)).ToList();

            //var selecionaLinha = listaValores.Select(linha => linha[18].Replace(",", "."));

            //foreach (var item in selecionaLinha)
            //{
            //    if (!string.IsNullOrEmpty(item))
            //        valorTotal += Convert.ToDecimal(item, CultureInfo.InvariantCulture);
            //}

            //var docsEncontrados = linhas.Where(linha => linha[34].Equals(documento)).Count();

            try
            {
                Parallel.ForEach(manutencoes, manutencao =>
                {
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
                        Quantidade_Orto = manutencao.Quantidade_Orto,
                        Tipo_Pagamento = manutencao.Tipo_Pagamento,
                        Vencimento = manutencao.Vencimento.ToString(),
                        Valor_Devido = manutencao.Valor_Devido.ToString()   ,
                        Valor_Total = manutencao.Valor_Total,
                        Data_Atendimento = manutencao.Data_Atendimento.ToString()
                    };

                    lstManutencoesDTO.Add(lstManutencao);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Manutencao: {error.Message}");
            }

            return lstManutencoesDTO;
        }

        public static List<ProcedimentosDTO> ConvertProcedimentosParaProcedimentosDTO(List<Procedimentos> procedimentos)
        {
            List<ProcedimentosDTO> lstProcedimentosDTO = new List<ProcedimentosDTO>();

            try
            {
                Parallel.ForEach(procedimentos, procedimento =>
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
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos: {error.Message}");
            }

            return lstProcedimentosDTO;
        }

        public static List<FinanceiroRecebiveisDTO> ConvertRecebiveisParaRecebiveisDTO(List<Recebivel> recebiveis)
        {
            List<FinanceiroRecebiveisDTO> lstRecebiveisDTO = new List<FinanceiroRecebiveisDTO>();

            try
            {
                Parallel.ForEach(recebiveis, recebivel =>
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
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos: {error.Message}");
            }

            return lstRecebiveisDTO;
        }

        public static List<FinanceiroRecebidosDTO> ConvertRecebidosParaRecebidosDTO(List<Recebidos> receber)
        {
            List<FinanceiroRecebidosDTO> lstReceberDTO = new List<FinanceiroRecebidosDTO>();

            try
            {
                Parallel.ForEach(receber, receber =>
                {
                    var lstReceber = new FinanceiroRecebidosDTO
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
                        Tipo_Pagamento = receber.Tipo_Pagamento,
                        Valor_Original = receber.Valor_Original.ToString(),
                        Vencimento_Recebivel = receber.Vencimento_Recebivel.ToShortDateString(),
                        Duplicata = receber.Duplicata,
                        Parcela = receber.Parcela.ToString(),
                        Tipo_Especie_Pagamento = receber.Tipo_Especie,
                        Especie_Pagamento = ExcelHelper.GetEspecieIDFromFormaPagamento(receber.Tipo_Especie).ToString()
                    };

                    lstReceberDTO.Add(lstReceber);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos: {error.Message}");
            }

            return lstReceberDTO;
        }

        public static List<AgendamentosDTO> ConvertAgendamentodsParaAgendamentosDTO(List<Agendamentos> agendamentos)
        {
            List<AgendamentosDTO> lstAgendamentosDTO = new List<AgendamentosDTO>();

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
                    //var idsEncontrados = agendamentos.Where(agenda => agenda.Equals(agendamento.ID)).Count();
                    //if (idsEncontrados > 0)
                    //    dataTermino = dataTermino.AddMinutes(15 * idsEncontrados);
                    //else
                    //    dataTermino = dataTermino.AddMinutes(15);

                    var lstAgendamento = new AgendamentosDTO
                    {
                        ID = agendamento.ID,
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
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Agendamentos: {error.Message}");
            }

            return lstAgendamentosDTO;
        }

        public static List<DesenvolvimentoClinicoDTO> ConvertDesenvolvimentoClinicoParaDesenvolvimentoClinicoDTO(List<DesenvolvimentoClinico> desenvClicnico)
        {
            List<DesenvolvimentoClinicoDTO> lstDesenvolvimentoClinicoDTO = new List<DesenvolvimentoClinicoDTO>();

            try
            {
                Parallel.ForEach(desenvClicnico, desenvClicnico =>
                {
                    var lstDesenvolvimentoClinico = new DesenvolvimentoClinicoDTO
                    {
                        Paciente_CPF = desenvClicnico.Paciente_CPF,
                        Paciente_Nome = desenvClicnico.Paciente_Nome,
                        Dentista_Nome = desenvClicnico.Dentista_Nome,
                        Dentista_Codigo = desenvClicnico.Dentista_Codigo,
                        Procedimento_Nome = desenvClicnico.Procedimento_Nome,
                        Data_Atendimento = desenvClicnico.Data_Retorno.ToString(),
                        Data_Inicio = desenvClicnico.Data_Inicio.ToString(),
                        Data_Retorno = desenvClicnico.Data_Retorno.ToString(),
                        Procedimento_Observacao = desenvClicnico.Procedimento_Observacao,
                        Lancamento = desenvClicnico.Lancamento
                    };

                    lstDesenvolvimentoClinicoDTO.Add(lstDesenvolvimentoClinico);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Desenvolvimento Clinico: {error.Message}");
            }

            return lstDesenvolvimentoClinicoDTO;
        }

        public static List<ProcedimentosManutencaoDTO> ConvertProcedManutParaProcedManutDTO(List<Procedimentos> procedimentos, List<Manutencoes> manutencoes)
        {
            List<ProcedimentosManutencaoDTO> lstProcedManutDTO = new List<ProcedimentosManutencaoDTO>();

            try
            {
                Parallel.ForEach(procedimentos, procedimento =>
                {
                    var lstProcedManut = new ProcedimentosManutencaoDTO
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

                    lstProcedManutDTO.Add(lstProcedManut);
                });

                Parallel.ForEach(manutencoes, manutencao =>
                {
                    var lstManutencao = new ProcedimentosManutencaoDTO
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
                        Quantidade_Orto = manutencao.Quantidade_Orto,
                        Tipo_Pagamento = manutencao.Tipo_Pagamento,
                        Vencimento = manutencao.Vencimento.ToString(),
                        Valor_Devido = manutencao.Valor_Devido.ToString(),
                        Valor_Total = manutencao.Valor_Total,
                        Data_Atendimento = manutencao.Data_Atendimento.ToString()
                    };

                    lstProcedManutDTO.Add(lstManutencao);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Procedimentos e Manutenções: {error.Message}");
            }

            return lstProcedManutDTO;
        }

        public static List<DentistasDTO> ConvertDentistasParaDentistasDTO(List<Dentistas> dentistas)
        {
            List<DentistasDTO> lstDentistasDTO = new List<DentistasDTO>();

            try
            {
                Parallel.ForEach(dentistas, dentista =>
                {
                    var lstDentistas = new DentistasDTO
                    {
                        Codigo = dentista.Codigo?.ToString(),
                        Ativo = dentista.Ativo.ToString(),
                        Nome_Completo = dentista.Nome_Completo,
                        NomeSocial = string.Empty,
                        Apelido = dentista.NOME?.GetPrimeirosCaracteres(20).ToNome(),
                        Data_Cadastro = dentista.Data_Cadastro.ToString(),
                        Observacoes = dentista.Observacoes,
                        Email = dentista.Email?.ToEmail(),
                        Nascimento_Local = string.Empty,
                        Estado_Civil = string.Empty,
                        Profissao = string.Empty,
                        Cargo_Clinica = string.Empty,
                        Dentista = "N",
                        Conselho_Codigo = string.Empty,
                        Paciente = "N",
                        Funcionario = "S",
                        Fornecedor = "N"
                    };

                    lstDentistasDTO.Add(lstDentistas);
                });
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
                Parallel.ForEach(recHistVendas, recHistVenda =>
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
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Recebiveis Histórico Venda: {error.Message}");
            }

            return lstRecebiveisHistVendaDTO;
        }

        public static List<GruposProcedimentosDTO> ConvertGruposProcedimentosParaGruposProcedimentosDTO(List<GruposProcedimentos> gruposProcedimentos)
        {
            List<GruposProcedimentosDTO> lstGruposProcedimentosDTO = new List<GruposProcedimentosDTO>();

            try
            {
                Parallel.ForEach(gruposProcedimentos, grupoProcedimento =>
                {
                    var lstGruposProcedimentos = new GruposProcedimentosDTO
                    {
                        NomeTabela = grupoProcedimento.Procedimento_Nome,
                        Especialidade = grupoProcedimento.Nome_Grupo,
                        Ativo = grupoProcedimento.Ativo,
                        NomeProcedimento = grupoProcedimento.Procedimento_Nome,
                        Abreviacao = grupoProcedimento.Abreviacao,
                        Preco = grupoProcedimento.Preco.ToString(),
                        TUSS = grupoProcedimento.TUSS,
                        EspecialidadeCodigo = grupoProcedimento.Codigo_Grupo
                    };

                    lstGruposProcedimentosDTO.Add(lstGruposProcedimentos);
                });
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
            }

            return lstGruposProcedimentosDTO;
        }

    }
}
