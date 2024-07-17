using Migracao.Models;
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

            //string arquivoDDentistasSql = "Scripts\\SelectDentistas.sql";

            //var dentistasClinico = new FireBirdContext<Dentistas>(_pathDB).RetornaItensBancoPorQuery(arquivoDDentistasSql);
            //var dentistasContrato = new FireBirdContext<Dentistas>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoDDentistasSql);
            //var dentistas = dentistasClinico.Union(dentistasContrato).ToList();

            var lstPacientesDentistas = ConversorEntidadeParaDTO.ConvertPacientesDentistasParaPacientesDentistasDTO(pacientes);

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
            var desenvClinicoContratos = new FireBirdContext<DesenvolvimentoClinico>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoDesenvClinicoSql);
            var desenvClinicoMerge = desenvClinico.Union(desenvClinicoContratos).ToList();

            string arquivoAgendamentosSql = "Scripts\\SelectAgendamentos.sql";
            var agendamentosClinico = new FireBirdContext<Agendamentos>(_pathDB).RetornaItensBancoPorQuery(arquivoAgendamentosSql);
            var agendamentosContratos = new FireBirdContext<Agendamentos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoDesenvClinicoSql);
            var agendamentosMerge = agendamentosClinico.Union(agendamentosContratos).ToList();

            var lstDesenvClinico = ConversorEntidadeParaDTO.ConvertDesenvolvimentoClinicoParaDesenvolvimentoClinicoDTO(desenvClinicoMerge, agendamentosMerge);

            var dataTableDesenvClinico = ExcelHelper.ConversorEntidadeParaDataTable(lstDesenvClinico);

            if (dataTableDesenvClinico != null)
            {
                var salvarDesenvClinico = Tools.GerarNomeArquivo($"CadastroDesenvClinico_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarDesenvClinico + ".xlsx", dataTableDesenvClinico);
            }
        }        

        void IDataBaseMigracao.DataBaseImportacaoManutencoes()
        {
            var excelHelper = new ExcelHelper();

            var dataTableProcedimentosManutencaoMerge = new DataTable();

            string arquivoSql = "Scripts\\SelectManutencoes.sql";
            var manutencoesClinico = new FireBirdContext<Manutencoes>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);
            var manutencoesContratos = new FireBirdContext<Manutencoes>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoSql);
            var manutencoesMerge = manutencoesClinico.Union(manutencoesContratos).ToList();

            var lstManutencoes = ConversorEntidadeParaDTO.ConvertManutencoesParaManutencoesDTO(manutencoesMerge);

            var dataTableManutencoes = ExcelHelper.ConversorEntidadeParaDataTable(lstManutencoes);

            if (dataTableManutencoes != null)
            {
                var salvarManutencoes = Tools.GerarNomeArquivo($"CadastroManutenções_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivoV2(salvarManutencoes + ".xlsx", dataTableManutencoes);
            }

            //string arquivoProdedimentosSql = "Scripts\\SelectProdedimentos.sql";
            //var procedimentosClicico = new FireBirdContext<Models.Procedimentos>(_pathDB).RetornaItensBancoPorQuery(arquivoProdedimentosSql);
            //var procedimentosContratos = new FireBirdContext<Procedimentos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoProdedimentosSql);
            //var procedimentos = procedimentosClicico.Union(procedimentosContratos).ToList();

            //var lstProcedimentos = ConversorEntidadeParaDTO.ConvertProcedimentosParaProcedimentosDTO(procedimentos);

            //var dataTableProcedimentos = ExcelHelper.ConversorEntidadeParaDataTable(lstProcedimentos);
            
            //if (dataTableProcedimentosManutencaoMerge != null)
            //{
            //    var salvarProcedimentos = Tools.GerarNomeArquivo($"cadastroProcedimentos_{Tools.ultimoEstabelecimentoID}_odontocompany");
            //    excelHelper.CriarExcelArquivo(salvarProcedimentos + ".xlsx", dataTableProcedimentos);
            //}
        }

        void IDataBaseMigracao.DataBaseImportacaoPagosExigiveis()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoSql = "Scripts\\SelectFinanceiroRecebidos.sql";
            var recebidosClinico = new FireBirdContext<Models.Recebidos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);
            var recebidosContratos = new FireBirdContext<Models.Recebidos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoSql);
            var recebidosMerge = recebidosClinico.Union(recebidosContratos).ToList();

            var lstRecebidos = ConversorEntidadeParaDTO.ConvertRecebidosParaRecebidosDTO(recebidosMerge);

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
            var procedimentosPrecosClinico = new FireBirdContext<Models.ProcedimentosPrecos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);
            var procedimentosPrecosContratos = new FireBirdContext<Models.ProcedimentosPrecos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoSql);
            var procedimentosPrecosMerge = procedimentosPrecosClinico.Union(procedimentosPrecosContratos).ToList();

            var lstProcedimentosPrecos = ConversorEntidadeParaDTO.ConvertProcedimentosPrecosParaProcedimentosPrecosDTO(procedimentosPrecosMerge);

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
