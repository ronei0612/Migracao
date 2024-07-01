using Migracao.DTO;
using Migracao.Models;
using Migracao.Models.DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Utils
{
    public class ConversorEntidadeParaDTO
    {
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

        public static List<FinanceiroRecebiveisDTO> ConvertRecebivelParaRecebivelDTO(List<Recebiveis> recebiveis)
        {
            List<FinanceiroRecebiveisDTO> lstRecebiveisDTO = new List<FinanceiroRecebiveisDTO>();

            try
            {
                Parallel.ForEach(recebiveis, recebivel =>
                {
                    var lstRecebiveis = new FinanceiroRecebiveisDTO
                    {
                        
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

        public static List<FinanceiroReceberDTO> ConvertRecebivelsParaRecebivelDTO(List<Receber> receber)
        {
            List<FinanceiroReceberDTO> lstReceberDTO = new List<FinanceiroReceberDTO>();

            try
            {
                Parallel.ForEach(receber, receber =>
                {
                    var lstReceber = new FinanceiroReceberDTO
                    {
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
    }
}
