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
            var excelHelper = new ExcelHelper();

            try
            {
                var procedimentosAgrupados = lstGruposProcedimentos
                    .GroupBy(p => p.Especialidade.Trim()) // Agrupa os procedimentos pela especialidade para salvar os arquivos
                    .Select(g => new { Especialidade = g.Key, Procedimentos = g.ToList() })
                    .ToList();

                foreach (var grupo in procedimentosAgrupados)
                {
                    var fileName = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_TabelaDePreços_{grupo.Especialidade}");
                    var dataTable = ExcelHelper.ConversorEntidadeParaDataTable(grupo.Procedimentos);

                    if (dataTable != null)
                        excelHelper.CriarExcelArquivoV2(fileName + ".xlsx", dataTable);
                }
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Tabela de Preços: {error.Message}");
            }
        }

        //Pricipais
        void IDataBaseMigracao.DataBaseImportacaoPacientesDentistas()
        {
            var excelHelper = new ExcelHelper();

            var arquivoPacientesSql = "Scripts\\SelectPacientes.sql";

            var pacientesClinico = new FireBirdContext<Pacientes>(_pathDB).RetornaItensBancoPorQuery(arquivoPacientesSql);
            var pacientesContrato = new FireBirdContext<Pacientes>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoPacientesSql);
            var pacientes = pacientesClinico.Union(pacientesContrato).DistinctBy(p => p.CPF).ToList();

            var lstPacientesDentistas = ConversorEntidadeParaDTO.ConvertPacientesDentistasParaPacientesDentistasDTO(pacientes);

            var dataTablePacientesDentistas = ExcelHelper.ConversorEntidadeParaDataTable(lstPacientesDentistas);

            if (dataTablePacientesDentistas != null)
            {
                var salvarArquivoPacientesDentistas = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_Pacientes");
                excelHelper.CriarExcelArquivoV2(salvarArquivoPacientesDentistas + ".xlsx", dataTablePacientesDentistas);
            }
        }              

        void IDataBaseMigracao.DataBaseImportacaoDevClinico()
        {
            var excelHelper = new ExcelHelper();

            var arquivoDesenvClinicoSql = "Scripts\\SelectDesenvolvimentoClinico.sql";
            var desenvClinico = new FireBirdContext<DesenvolvimentoClinico>(_pathDB).RetornaItensBancoPorQuery(arquivoDesenvClinicoSql);
            var desenvClinicoContratos = new FireBirdContext<DesenvolvimentoClinico>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoDesenvClinicoSql);
            //var desenvClinicoMerge = desenvClinico.Union(desenvClinicoContratos).DistinctBy(x => x.Paciente_CPF).ToList();
            var desenvClinicoMerge = desenvClinico.Union(desenvClinicoContratos).ToList();

            var arquivoAgendamentosSql = "Scripts\\SelectAgendamentos.sql";
            var agendamentosClinico = new FireBirdContext<Agendamentos>(_pathDB).RetornaItensBancoPorQuery(arquivoAgendamentosSql);
            var agendamentosContratos = new FireBirdContext<Agendamentos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoAgendamentosSql);
            //var agendamentosMerge = agendamentosClinico.Union(agendamentosContratos).DistinctBy(x => x.Paciente_CPF).ToList();
            var agendamentosMerge = agendamentosClinico.Union(agendamentosContratos).ToList();

            var lstDesenvClinico = ConversorEntidadeParaDTO.ConvertDesenvolvimentoClinicoParaDesenvolvimentoClinicoDTO(desenvClinicoMerge, agendamentosMerge);

            var dataTableDesenvClinico = ExcelHelper.ConversorEntidadeParaDataTable(lstDesenvClinico);

            if (dataTableDesenvClinico != null)
            {
                var salvarDesenvClinico = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_DesenvClinico");
                excelHelper.CriarExcelArquivoV2(salvarDesenvClinico + ".xlsx", dataTableDesenvClinico);
            }
        }        

        void IDataBaseMigracao.DataBaseImportacaoManutencoes()
        {
            var excelHelper = new ExcelHelper();

            var dataTableProcedimentosManutencaoMerge = new DataTable();

            var arquivoSql = "Scripts\\SelectManutencoes.sql";
            var manutencoesClinico = new FireBirdContext<Manutencoes>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);
            var manutencoesContratos = new FireBirdContext<Manutencoes>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoSql);
            var manutencoesMerge = manutencoesClinico.Union(manutencoesContratos).ToList();

            var lstManutencoes = ConversorEntidadeParaDTO.ConvertManutencoesParaManutencoesDTO(manutencoesMerge);

            var dataTableManutencoes = ExcelHelper.ConversorEntidadeParaDataTable(lstManutencoes);

            if (dataTableManutencoes != null)
            {
                var salvarManutencoes = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_Manutenções");
                excelHelper.CriarExcelArquivoV2(salvarManutencoes + ".xlsx", dataTableManutencoes);
            }

            var arquivoProcedimentosSql = "Scripts\\SelectProcedimentos.sql";
            var procedimentosClicico = new FireBirdContext<Models.Procedimentos>(_pathDB).RetornaItensBancoPorQuery(arquivoProcedimentosSql);
            var procedimentosContratos = new FireBirdContext<Procedimentos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoProcedimentosSql);
            var procedimentos = procedimentosClicico.Union(procedimentosContratos).ToList();

            var lstProcedimentos = ConversorEntidadeParaDTO.ConvertProcedimentosParaProcedimentosDTO(procedimentos);

            var dataTableProcedimentos = ExcelHelper.ConversorEntidadeParaDataTable(lstProcedimentos);

            if (dataTableProcedimentosManutencaoMerge != null)
            {
                var salvarProcedimentos = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_Procedimentos");
                excelHelper.CriarExcelArquivoV2(salvarProcedimentos + ".xlsx", dataTableProcedimentos);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoPagosExigiveis()
        {
            var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectFinanceiroRecebidos.sql";

            var recebidosClinico = new FireBirdContext<Recebidos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);
            var recebidosContratos = new FireBirdContext<Recebidos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoSql);
            var recebidosMerge = recebidosClinico.Union(recebidosContratos).ToList();

            var lstRecebidos = ConversorEntidadeParaDTO.ConvertRecebidosParaRecebidosDTO(recebidosMerge);

            var dataTableRecebidos = ExcelHelper.ConversorEntidadeParaDataTable(lstRecebidos);

            if (dataTableRecebidos != null)
            {
                var salvarArquivoRecebidos = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_Financeiro");
                excelHelper.CriarExcelArquivoV2(salvarArquivoRecebidos + ".xlsx", dataTableRecebidos);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoProcedimentosPrecos()
        {
            //var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectProcedimentosPrecos.sql";
            var procedimentosPrecosClinico = new FireBirdContext<ProcedimentosPrecos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);
            var procedimentosPrecosContratos = new FireBirdContext<ProcedimentosPrecos>(_pathDBContratos).RetornaItensBancoPorQuery(arquivoSql);
            var procedimentosPrecosMerge = procedimentosPrecosClinico.Union(procedimentosPrecosContratos).ToList();

            var lstProcedimentosPrecos = ConversorEntidadeParaDTO.ConvertProcedimentosPrecosParaProcedimentosPrecosDTO(procedimentosPrecosMerge);

            RetornaProcedimentosPorTipoEntidade(lstProcedimentosPrecos);
        }


        // Complementares
        void IDataBaseMigracao.DataBaseImportacaoFinanceiroRecebiveis()
        {
            var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectFinanceiroRecebiveis.sql";

            var recebiveis = new FireBirdContext<Recebivel>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstRecebiveis = ConversorEntidadeParaDTO.ConvertRecebiveisParaRecebiveisDTO(recebiveis);

            var dataTableRecebiveis = ExcelHelper.ConversorEntidadeParaDataTable(lstRecebiveis);

            if (dataTableRecebiveis != null)
            {
                var salvarArquivoRecebiveis = Tools.GerarNomeArquivo($"CadastroRecebiveis_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivoV2(salvarArquivoRecebiveis + ".xlsx", dataTableRecebiveis);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoDentistas()
        {
            var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectDentistas.sql";

            var dentistas = new FireBirdContext<Models.Dentistas>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstDentistas = ConversorEntidadeParaDTO.ConvertDentistasParaDentistasDTO(dentistas);

            var dataTableDentistas = ExcelHelper.ConversorEntidadeParaDataTable(lstDentistas);

            if (dataTableDentistas != null)
            {
                var salvarArquivoDentistas = Tools.GerarNomeArquivo($"CadastroDentistas_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivoV2(salvarArquivoDentistas + ".xlsx", dataTableDentistas);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoRecebiveisHistVenda()
        {
            var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectRecebiveisHistVenda.sql";

            var recebiveisHistVenda = new FireBirdContext<Models.RecebiveisHistVenda>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstRecebiveisHistVenda = ConversorEntidadeParaDTO.ConvertRecebiveisHistVendaParaRecebiveisHistVendaDTO(recebiveisHistVenda);

            var dataTableRecebiveisHistVenda = ExcelHelper.ConversorEntidadeParaDataTable(lstRecebiveisHistVenda);

            if (dataTableRecebiveisHistVenda != null)
            {
                var salvarArquivoRecebiveisHistVenda = Tools.GerarNomeArquivo($"CadastroRecebiveisHistVenda_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivoV2(salvarArquivoRecebiveisHistVenda + ".xlsx", dataTableRecebiveisHistVenda);
            }
        }        

        void IDataBaseMigracao.DataBaseImportacaoProntuarios()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoAgendamentos()
        {
            var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectAgendamentos.sql";

            var agendamentos = new FireBirdContext<Agendamentos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstAgendamentos = ConversorEntidadeParaDTO.ConvertAgendamentodsParaAgendamentosDTO(agendamentos);

            var dataTableAgendamentos = ExcelHelper.ConversorEntidadeParaDataTable(lstAgendamentos);

            if (dataTableAgendamentos != null)
            {
                var salvarArquivoAgendamentos = Tools.GerarNomeArquivo($"CadastroAgendamentos_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivoV2(salvarArquivoAgendamentos + ".xlsx", dataTableAgendamentos);
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoProcedimentos()
        {
            var excelHelper = new ExcelHelper();

            var arquivoSql = "Scripts\\SelectProdedimentos.sql";

            var procedimentos = new FireBirdContext<Models.Procedimentos>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

            var lstProcedimentos = ConversorEntidadeParaDTO.ConvertProcedimentosParaProcedimentosDTO(procedimentos);

            var dataTableProcedimentos = ExcelHelper.ConversorEntidadeParaDataTable(lstProcedimentos);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentos = Tools.GerarNomeArquivo($"{Tools.ultimoEstabelecimentoID}_OdontoCompany_Cadastro_Procedimentos");
                excelHelper.CriarExcelArquivoV2(salvarProcedimentos + ".xlsx", dataTableProcedimentos);
            }
        }
    }
}
