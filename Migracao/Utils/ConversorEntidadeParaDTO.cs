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
                        NomeCompleto = paciente.NOME,
                        NomeSocial = string.Empty,
                        Apelido = paciente.NOME.GetPrimeirosCaracteres(20).ToNome(),
                        Documento = paciente.CGC_CPF.ToCPF(),
                        DataCadastro = paciente.DT_CADASTRO,
                        Observacoes = paciente.OBS1,
                        Email = paciente.EMAIL,
                        RG = paciente.INSC_RG.GetPrimeirosCaracteres(20),
                        Sexo = paciente.SEXO_M_F.ToSexo("m", "f") ? "M" : "F",
                        NascimentoData = paciente.DT_NASCIMENTO,
                        Paciente = paciente.CLIENTE,
                        Funcionario = "N",
                        Fornecedor = paciente.FORNECEDOR,
                        TelefonePrincipal = paciente.FONE1,
                        Celular = paciente.CELULAR,
                        TelefoneAlternativo = paciente.FONE2,
                        Logradouro = paciente.ENDERECO.PrimeiraLetraMaiuscula(),
                        LogradouroNum = paciente.NUM_ENDERECO,
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
                        NumeroControle = manutencao.NumeroControle,
                        PacienteCPF = manutencao.PacienteCPF,
                        PacienteNome = manutencao.PacienteNome,
                        DentistaNome = manutencao.DentistaNome,
                        ProcedimentoNome = manutencao.ProcedimentoNome,
                        ProcedimentoValor = manutencao.ProcedimentoValor,
                        ValorOriginal = manutencao.ValorOriginal,
                        ValorPagamento = manutencao.ValorPagamento,
                        DataPagamento = manutencao.DataPagamento,
                        Dente = manutencao.Dente,
                        ProcedimentoObservacao = manutencao.ProcedimentoObservacao,
                        QuantidadeOrto = manutencao.QuantidadeOrto,
                        TipoPagamento = manutencao.TipoPagamento,
                        Vencimento = manutencao.Vencimento,
                        ValorDevido = manutencao.ValorDevido,
                        ValorTotal = manutencao.ValorTotal,
                        DataAtendimento = manutencao.DataAtendimento
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
                        NumeroControle = procedimento.Numero_Controle,
                        PacienteCPF = procedimento.Paciente_CPF.ToCPF(),
                        PacienteNome = procedimento.Nome_Paciente,
                        DentistaCPF = procedimento.Dentista_CPF,
                        DentistaNome = procedimento.Dentista_Nome,
                        Dente = procedimento.Dente,
                        ProcedimentoNome = procedimento.NOME_PRODUTO,
                        ProcedimentoValor = procedimento.Valor,
                        ProcedimentoObservacao = procedimento.Observacao,
                        DataInicio = procedimento.Data_Inicio.ToData().ToShortDateString(),
                        DataTermino = procedimento.Data_Termino,
                        DataAtendimento = procedimento.Data_Atendimento
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


        // TODO
        // Prontuários e Desenvolvimento Clinico
    }
}
