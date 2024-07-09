﻿using Migracao.Models;
using Migracao.Models.Context;
using Migracao.Models.DTO;
using Migracao.Models.Interfaces;
using Migracao.Utils;
using NPOI.SS.Formula.Functions;
using System.Data;
using System.Linq;
using System.Text;

namespace Migracao.Sistems
{
    internal class OdontoCompany : IDataBaseMigracao
    {
        FireBirdContext<T> context;
        string _pathDB;
        string _pathDBContratos;

        public OdontoCompany(string dataBaseName, string pathDB = null, string pathDBContratos = null)
        {
            _pathDB = pathDB;
            _pathDBContratos = pathDBContratos;
            var bla = new FireBirdContext<Models.OdontoCompany.Atendimento222>(pathDB).GetAll();
        }

        public Tuple<List<string[]>, List<string>> LerArquivosExcelCsv(string arquivo, Encoding encoding)
        {
            var separador = ExcelHelper.DetectarSeparadorCSV(arquivo);

            List<string> cabecalhosCSV = ExcelHelper.GetCabecalhosCSV(arquivo, separador, encoding);
            List<string[]> linhasCSV = ExcelHelper.GetLinhasCSV(arquivo, separador, cabecalhosCSV.Count(), encoding);

            return new Tuple<List<string[]>, List<string>>(linhasCSV, cabecalhosCSV);
        }

        public void RetornaProcedimentosPorTipoEntidade(List<ProcedimentosPrecosDTO> lstGruposProcedimentos)
        {
            ExcelHelper excelHelper = new();

            try
            {
                var procedimentosAgrupados = lstGruposProcedimentos
                    .GroupBy(p => p.Especialidade.Trim()) // Agrupa os procedimentos pela especialidade para salvar os arquivos
                    .Select(g => new { Especialidade = g.Key, Procedimentos = g.ToList() })
                    .ToList();

                foreach (var grupo in procedimentosAgrupados)
                {
                    var fileName = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_CadastroProcedimentos_{grupo.Especialidade}");
                    var dataTable = ExcelHelper.ConversorEntidadeParaDataTable(grupo.Procedimentos);

                    if (dataTable != null)
                        excelHelper.CriarExcelArquivo(fileName + ".xlsx", dataTable);
                }
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Grupo Procedimentos: {error.Message}");
            }
        }

        //Pricipais
        void IDataBaseMigracao.DataBaseImportacaoPacientesDentistas()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoPacientesSql = "Scripts\\SelectPacientes.sql";

            var pacientesClinico = new FireBirdContext<Pacientes>(_pathDB).RetornaItensBancoPorQuery(arquivoPacientesSql);
            var pacientesContrato = new FireBirdContext<Pacientes>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoPacientesSql);
            var pacientes = pacientesClinico.Union(pacientesContrato).ToList();

            string arquivoDDentistasSql = "Scripts\\SelectDentistas.sql";

            var dentistasClinico = new FireBirdContext<Dentistas>(_pathDB).RetornaItensBancoPorQuery(arquivoDDentistasSql);
            var dentistasContrato = new FireBirdContext<Dentistas>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoDDentistasSql);
            var dentistas = dentistasClinico.Union(dentistasContrato).ToList();

            var lstPacientesDentistas = ConversorEntidadeParaDTO.ConvertPacientesDentistasParaPacientesDentistasDTO(pacientes, dentistas);

            var dataTablePacientesDentistas = ExcelHelper.ConversorEntidadeParaDataTable(lstPacientesDentistas);

            if (dataTablePacientesDentistas != null)
            {
                var salvarArquivoPacientesDentistas = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_CadastroPacientesDentistas");
                excelHelper.CriarExcelArquivo(salvarArquivoPacientesDentistas + ".xlsx", dataTablePacientesDentistas);
            }
        }              

        void IDataBaseMigracao.DataBaseImportacaoDevClinico()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoDesenvClinicoSql = "Scripts\\SelectDesenvolvimentoClinico.sql";

            var desenvClinico = new FireBirdContext<DesenvolvimentoClinico>(_pathDB).RetornaItensBancoPorQuery(arquivoDesenvClinicoSql);

            string arquivoAgendamentosSql = "Scripts\\SelectAgendamentos.sql";

            var agendamentos = new FireBirdContext<Agendamentos>(_pathDB).RetornaItensBancoPorQuery(arquivoAgendamentosSql);

            var lstDesenvClinico = ConversorEntidadeParaDTO.ConvertDesenvolvimentoClinicoParaDesenvolvimentoClinicoDTO(desenvClinico, agendamentos);

            var dataTableDesenvClinico = ExcelHelper.ConversorEntidadeParaDataTable(lstDesenvClinico);

            if (dataTableDesenvClinico != null)
            {
                var salvarDesenvClinico = Tools.GerarNomeArquivo($"CadastroDesenvClinico_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarDesenvClinico + ".xlsx", dataTableDesenvClinico);
            }
        }        

        void IDataBaseMigracao.DataBaseImportacaoManutencoes()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            DataTable dataTableProcedimentosManutencaoMerge = new();

            string arquivoSql = "Scripts\\SelectManutencoes.sql";

            var manutencoes = new FireBirdContext<Manutencoes>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstManutencoes = ConversorEntidadeParaDTO.ConvertManutencoesParaManutencoesDTO(manutencoes);

            var dataTableManutencoes = ExcelHelper.ConversorEntidadeParaDataTable(lstManutencoes);

            if (dataTableManutencoes != null)
            {
                var salvarManutencoes = Tools.GerarNomeArquivo($"CadastroManutenções_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarManutencoes + ".xlsx", dataTableManutencoes);
            }

            string arquivoProdedimentosSql = "Scripts\\SelectProdedimentos.sql";

            var procedimentos = new FireBirdContext<Models.Procedimentos>(_pathDB).RetornaItensBancoPorQuery(arquivoProdedimentosSql);

            var lstProcedManut = ConversorEntidadeParaDTO.ConvertProcedManutParaProcedManutDTO(procedimentos, manutencoes);

            var dataTableProcedManut = ExcelHelper.ConversorEntidadeParaDataTable(lstProcedManut);
            
            if (dataTableProcedimentosManutencaoMerge != null)
            {
                var salvarProcedimentosManutencaoMerge = Tools.GerarNomeArquivo($"cadastroProcedimentosManutencaoEntidadesMerge_{Tools.ultimoEstabelecimentoID}_odontocompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosManutencaoMerge + ".xlsx", dataTableProcedManut);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoPagosExigiveis()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectFinanceiroRecebidos.sql";

            var recebidos = new FireBirdContext<Models.Recebidos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstRecebidos = ConversorEntidadeParaDTO.ConvertRecebidosParaRecebidosDTO(recebidos);

            var dataTableRecebidos = ExcelHelper.ConversorEntidadeParaDataTable(lstRecebidos);

            if (dataTableRecebidos != null)
            {
                var salvarArquivoRecebidos = Tools.GerarNomeArquivo($"CadastroRecebidos_Baixas_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoRecebidos + ".xlsx", dataTableRecebidos);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoProcedimentosPrecos()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            int estabelecimentoID = 1;

            string arquivoSql = "Scripts\\SelectProcedimentosPrecos.sql";

            var procedimentosPrecos = new FireBirdContext<Models.ProcedimentosPrecos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstProcedimentosPrecos = ConversorEntidadeParaDTO.ConvertProcedimentosPrecosParaProcedimentosPrecosDTO(procedimentosPrecos);

            RetornaProcedimentosPorTipoEntidade(lstProcedimentosPrecos);
        }


        // Complementares
        void IDataBaseMigracao.DataBaseImportacaoFinanceiroRecebiveis()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectFinanceiroRecebiveis.sql";

            var recebiveis = new FireBirdContext<Models.Recebivel>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstRecebiveis = ConversorEntidadeParaDTO.ConvertRecebiveisParaRecebiveisDTO(recebiveis);

            var dataTableRecebiveis = ExcelHelper.ConversorEntidadeParaDataTable(lstRecebiveis);

            if (dataTableRecebiveis != null)
            {
                var salvarArquivoRecebiveis = Tools.GerarNomeArquivo($"CadastroRecebiveis_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoRecebiveis + ".xlsx", dataTableRecebiveis);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoDentistas()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectDentistas.sql";

            var dentistas = new FireBirdContext<Models.Dentistas>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstDentistas = ConversorEntidadeParaDTO.ConvertDentistasParaDentistasDTO(dentistas);

            var dataTableDentistas = ExcelHelper.ConversorEntidadeParaDataTable(lstDentistas);

            if (dataTableDentistas != null)
            {
                var salvarArquivoDentistas = Tools.GerarNomeArquivo($"CadastroDentistas_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoDentistas + ".xlsx", dataTableDentistas);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoRecebiveisHistVenda()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectRecebiveisHistVenda.sql";

            var recebiveisHistVenda = new FireBirdContext<Models.RecebiveisHistVenda>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstRecebiveisHistVenda = ConversorEntidadeParaDTO.ConvertRecebiveisHistVendaParaRecebiveisHistVendaDTO(recebiveisHistVenda);

            var dataTableRecebiveisHistVenda = ExcelHelper.ConversorEntidadeParaDataTable(lstRecebiveisHistVenda);

            if (dataTableRecebiveisHistVenda != null)
            {
                var salvarArquivoRecebiveisHistVenda = Tools.GerarNomeArquivo($"CadastroRecebiveisHistVenda_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoRecebiveisHistVenda + ".xlsx", dataTableRecebiveisHistVenda);
            }
        }        

        void IDataBaseMigracao.DataBaseImportacaoProntuarios()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoAgendamentos()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectAgendamentos.sql";

            var agendamentos = new FireBirdContext<Agendamentos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstAgendamentos = ConversorEntidadeParaDTO.ConvertAgendamentodsParaAgendamentosDTO(agendamentos);

            var dataTableAgendamentos = ExcelHelper.ConversorEntidadeParaDataTable(lstAgendamentos);

            if (dataTableAgendamentos != null)
            {
                var salvarArquivoAgendamentos = Tools.GerarNomeArquivo($"CadastroAgendamentos_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoAgendamentos + ".xlsx", dataTableAgendamentos);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoProcedimentos()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectProdedimentos.sql";

            var procedimentos = new FireBirdContext<Models.Procedimentos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstProcedimentos = ConversorEntidadeParaDTO.ConvertProcedimentosParaProcedimentosDTO(procedimentos);

            var dataTableProcedimentos = ExcelHelper.ConversorEntidadeParaDataTable(lstProcedimentos);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentos = Tools.GerarNomeArquivo($"CadastroProcedimentos_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentos + ".xlsx", dataTableProcedimentos);
            }
        }
    }
}
