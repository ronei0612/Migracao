using EnumsNET;
using MathNet.Numerics;
using MathNet.Numerics.Distributions;
using Migracao.Models;
using Migracao.Models.Context;
using Migracao.Models.DentalOffice;
using Migracao.Models.DTO;
using Migracao.Models.Interfaces;
using Migracao.Models.OdontoCompany;
using Migracao.Utils;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using System.Data;
using System.Globalization;
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

        string arquivoExcelNomesUTF8 = "Files\\NomesUTF8.xlsx";

        string[] EMD101_Pacientes = ["BAIRRO", "CEP", "CGC_CPF", "CIDADE", "CLIENTE", "CELULAR", "DT_CADASTRO", "DT_NASCIMENTO", "EMAIL", "ENDERECO", "ESTADO", "FONE1", "FONE2", "FORNECEDOR", "INSC_RG", "NOME", "NUM_CONVENIO", "NUM_ENDERECO", "NUM_FICHA", "OBS1", "SEXO_M_F"];
        string[] CRD111_Recebiveis = ["AGENCIA", "AGUARDANDO_VINCULO", "ALINEA", "AXON_ID", "BANCO", "BANDA1", "BANDA2", "BANDA3", "BAIXA", "CAMPOX", "COD_CAIXA", "CODIGO_TUSS", "CONTA", "CONTA_CORRENTE", "CONTA_DOCUMENTO", "COBRADORA", "COBRANCA", "DATA_ENV_CART", "DATA_ENV_SCPC", "DATA_REMESSA", "DATA_RET_CART", "DATA_RET_SCPC", "DESCONTO_BOLETO", "DEVOLUCAO", "DOCUMENTO", "DT_AXON", "DUPLICATA", "EMISSAO", "EMITENTE", "ENCARGOS", "FILIAL", "GEROU_TRANSMISSAO", "GRUPO", "ID_BAIXAPLANOS", "ID_PIX", "JUROS", "LANCTO", "LOCAL", "LOJA", "MODIFICADO", "MOTIVO", "MULTA", "NOME_GRUPO", "NOME_LOCAL", "NOSSONUMERO", "NUM_BANCO", "OBS", "ORDEM", "PARCELA", "PERIODO", "PRAZO", "REAPRESENTOU", "RECEBEU_TRANSMISSAO", "REMESA", "RESPONSAVEL", "SEQ_ALINEA11", "SITUACAO", "SITUACAO_REMESSA", "TERMINAL", "TIPO_COBRANCA", "TIPO_DOC", "TOTAL", "TRANSMISSAO", "USUARIO", "VALOR", "VALOR_ORIG", "VALOR_RECEBER", "VALOR_VENDA", "VENCTO", "VENCTO_ORIG", "VR_CALCULADO", "VR_PARCELA"];
        string[] CXD555_Baixa = ["AGENCIA", "BANCO", "BAIXA", "CALCULO", "CNPJ_CPF", "CONTA", "DATA", "DOCUMENTO", "DT_AXON", "DT_DEPOSITO", "FECHAR_DIRETO", "FICHA_FINANCEIRO", "HISTORICO", "HORA", "LANCTO", "LOJA", "LOTE", "MODIFICADO", "NUM_CONVENIO", "OBS1", "OBS2", "OBS3", "PERIODO", "PRO_MED", "PRO_ODO", "RESPONSAVEL", "ROY_MED", "ROY_ODO", "TERMINAL", "TIPO", "TRANSMISSAO", "USUARIO", "VALOR", "VALOR_RECEBER", "VLR_BRUTO"];
        string[] BXD111_Baixa = ["AGUARDANDO_VINCULO", "AXON_ID", "BAIXA", "BANCO", "CAMPOX", "CGC_CPF", "COD_CAIXA", "CONTA_CORRENTE", "CONTA_DOCUMENTO", "DATA_REMESSA", "DOCUMENTO", "DT_AXON", "DUPLICATA", "GRUPO", "ID_BAIXAPLANOS", "LANCTO", "LOJA", "MODIFICADO", "MOTIVO", "NOME_GRUPO", "NUM_BANCO", "OBS", "PARCELA", "RESPONSAVEL", "TERMINAL", "TIPO_DOC", "TRANSMISSAO", "USUARIO", "VALOR", "VENCTO", "VR_CALCULADO", "VR_PARCELA"];
        string[] CRD013_Especies = ["CODIGO", "NOME", "CAMPOX", "HISTORICO", "GRAVAR_CAIXA", "GRAVAR_DUPLICATA", "BAIXAR_AUTOMATICO", "DIAS_PRAZO", "CONTA", "SCUSTO", "BAIXAR_GRAVAR_CAIXA", "CARENCIA", "PERC_ACRESCIMO", "PERC_DESCONTO", "PERC_JUROS", "MULTA", "QTDE_PARCELAS", "CODIGO_CAIXA", "PRO_RATA", "TAXA_CARTAO", "RESUMIR_BOBINA", "TARIFA", "RESUMIR_F767", "ECF", "DIAS_CC", "CODIGO_ENTRADA", "CARNE", "DUPLICATA", "INICIAIS_FXD111", "GRAVAR_CC_NA_EMISSAO", "CODIGO_RETIRADA", "ENCARGOS", "PERC_COMISSAO", "DESCONTO_MAXIMO_VENDA", "LOJA", "USUARIO", "MODIFICADO", "ATIVO", "OBRIGAR_DIAS_ADICIONAIS", "MAX_DIAS_PRIMEIRA", "ACRESCIMO_MAXIMO_VENDA", "OBRIGAR_ACRESCIMO", "OBRIGAR_DESCONTO", "OBRIGAR_ENTRADA", "A_VISTA", "NOME_BOBINA", "GRAVAR_PAGAR", "GRUPO	NOME_GRUPO", "BAIXAR_RECEBER", "BAIXAR_PAGAR", "GRAVAR_MANUTENCAO", "GRAVAR_ESTOQUE", "VER_CONVENIO", "FRANQUIA", "ORDEM", "BAIXA_BANCO", "MANUTENCAO_BANCO", "UNIPLAN", "VENDA_CONVENIO", "VLR_TITULAR", "VLR_DEPENDENTE", "COD_VENDA_ODC", "ORTODONTIA", "PARTMED", "BAIXA_COBRADORA", "TIPO_CONTRATO", "TAXA_ADESAO	RENOVACAO", "BAIXA_CARTAO", "BAIXA_ORTO", "CONSULTA_PARTMED", "QTDE_CONSULTAS", "BAIXAR_CODIGO", "GERAR_CHEQUE", "GERAR_TAXA_ADESAO", "ESTORNO_ROYALTIES", "FORMA_PAGAMENTO", "DT_AXON", "AXON_ID", "GRUPO_FORMA_PAGTO", "FORMA_PAGTO_CLINICORP"];
        string[] CED006_Dentistas = ["ADMISSAO", "AGENCIA", "AGENCIA2", "AXON_ID", "BAIRRO", "BANCO", "BANCO2", "CAIXA_POSTAL", "CEP", "CGC_CPF", "CIDADE", "CLIENTE", "CODIGO", "CODIGO_CLIENTE", "CODIGO_INDICACAO", "CODIGO_VALIDADE", "COD_MUNICIPIO", "COD_PRAMELHOR", "COD_UF", "COD_VENDEDOR", "CONJUGE", "CONTA", "CONTA2", "CPF_FIA", "CPF_INDICACAO", "DATA_APROVACAO_DRCASH", "DATA_BLOQUEIO", "DATA_DEP_EXCLUIDO", "DATA_LGPD", "DATA_VALIDADE", "DEPENDENTE", "DT_AXON", "DT_CADASTRO", "DT_NASC_FIA", "DT_NASCIMENTO", "DT_NASCIMENTO_DEP", "DT_ULTMOV", "EMAIL", "ENDERECO", "ENDERECO_FIA", "ESTADO", "ESTADO_FIA", "FAX", "FONE1", "FONE2", "FONE_FIA", "FONE_REF_1", "FONE_REF_2", "F_OU_J", "FORNECEDOR", "FUNCAO", "ID_DRCASH", "INSC_RG", "INSTITUTO_ODC", "LGPD_CPF", "LGPD_DATA_HORA", "LGPD_IMAGEM", "LGPD_MENSAGEM", "LGPD_TELEFONE", "LGPD_USUARIO", "LOJA", "MAE", "MODIFICADO", "NOME", "NOME_FIA", "NOME_GRUPO", "NOME_LOCAL", "NOME_REF_1", "NOME_REF_2", "NOME_VALIDADE", "NUM_BLOQUEIO", "NUM_CONVENIO", "NUM_ENDERECO", "NUM_FICHA", "OBS1", "OBS_VALIDADE", "ONDE_TRABALHA", "PAI", "PARENTESCO_FIA", "PRESTADOR", "PROFISSAO", "PROFISSAO_FIA", "PROTETICO", "PROTETICO_ATIVO", "QTDE_DEPENDENTES", "RENDA_FIA", "RENDA_MES", "RG_FIA", "SEXO_M_F", "TITULAR", "TITULAR_DEP_EXCLUIDO", "TRANSMISSAO", "USU_BLOQUEIO", "USU_CADASTRO", "USUARIO", "USUARIO_LGPD", "USUARIO_VALIDADE", "VALOR_MAXIMO_DRCASH", "VR_LIMITE"];
        string[] CED001_Procedimentos = ["CODIGO", "NOME", "SIMBOLO", "GRUPO", "CODIGO_TUSS", "VRVENDA", "ATIVO", "OBS", "PARTICULAR"];
        string[] CED002_CodProcedimentos = ["CODIGO", "NOME", "USUARIO"];
        string[] MAN001_Manutencao = ["CNPJ_CPF", "QTDE_MANUT", "OBS", "ATIVO", "INICIO_MANUT", "FINAL_MANUT", "MODIFICADO", "USUARIO", "DATA_ATIVO", "USU_ATIVO", "USU_INSTALACAO", "USU_RETIRADA", "MOTIVO_ATIVO", "DIAGNOSTICO", "PROGNOSTICO", "OBS_CLASSE", "CLASSE", "DATA_ALTERACAO", "HORA_ALTERACAO", "USUARIO_ALTERACAO", "DATA_MODIFICADO", "DT_AXON", "AXON_ID"];
        string[] MAN101_Manutencao = ["LANCTO", "NUM_MAN", "CNPJ_CPF", "CONTROLE", "DATA_PAGO", "DATA_RETORNO", "RESP_ATEND", "NOME_RESP_ATEND", "OBS_ATEND", "TIPO_ATEND", "DATA_LANC", "DATA_MANUT", "DATA_MODIFICADO"];
        string[] MAN111_Manutencao = ["LANCTO", "CNPJ_CPF", "DATA_PAGTO", "TIPO_PAGTO", "VALOR", "MES_ANO", "CAMPOX", "DIA_MES_ANO", "RESPONSAVEL", "DATA_ATEND", "RESP_ATEND", "NUM_MAN", "NOME_RESP_ATEND", "OBS", "RETORNO", "DATA_LANC", "HORA", "VENCTO", "VALOR_PARCELA", "NOSSO_NUMERO", "TIPO_MAN", "VENCTO_ORIG", "VALOR_ORIG", "MOTIVO_ALTERAR", "AUTORIZA_ALTERAR", "MOTIVO_INCLUIR", "AUTORIZA_INCLUIR", "VALOR_CALCULADO", "AUTORIZA_RECEBER", "MOTIVO_RECEBER", "EMISSAO", "BANCO", "DATA_ALTERACAO", "OPERACAO", "DOCUMENTO", "NOME_TIPO", "TIPO_COBRANCA", "TERMINAL", "REMESSA", "SENHA_ATEND", "USUARIO", "EM_ATENDIMENTO", "NOME_ATEND", "CALC_JUROS", "USU_VENDA", "COBRADORA", "COBRANCA", "DATA_REMESSA", "SITUACAO_REMESSA", "DATA_MODIFICADO", "CONTA", "NSU_TRANSACAO", "COBRANCA_ORTO", "DATA_REMESSA_CARTAO", "CONTROLE_CARTAO", "UNIDADE", "HORA_ATEND", "HORA_FINAL", "DESCONTO_BOLETO", "AGUARDANDO_VINCULO", "ID_PIX", "DT_AXON", "AXON_ID"];
        string[] ATD222_Procedimentos = ["DOCUMENTO", "CNPJ_CPF", "PRODUTO", "TIPO", "NOME_PRODUTO", "DATA", "VALOR", "OBS", "CAMPOX", "RESPONSAVEL", "DATA_ATEND", "RESP_ATEND", "NOME_RESP_ATEND", "NUMERO", "CONTROLE", "CBARRA", "TERMINAL", "USUARIO", "QTDE_SESSAO", "QTDE_SESSAO_ORIG", "MODIFICADO", "DATA_CANCELADO", "USUARIO_CANCELADO", "DATA_INCLUIDO", "USUARIO_INCLUIDO", "FORMA_MEDIDA", "QTDE_MEDIDA", "DT_AXON", "AXON_ID"];

        List<string> cabecalhos_Pacientes = ["Código", "Ativo(S/N)", "NomeCompleto", "NomeSocial", "Apelido", "Documento(CPF,CNPJ,CGC)", "DataCadastro(01/12/2024)", "Observações", "Email", "RG", "Sexo(M/F)", "NascimentoData", "NascimentoLocal", "EstadoCivil(S/C/V)", "Profissao", "CargoNaClinica", "Dentista(S/N)", "ConselhoCodigo", "Paciente(S/N)", "Funcionario(S/N)", "Fornecedor(S/N)", "TelefonePrincipal", "Celular", "TelefoneAlternativo", "Logradouro", "LogradouroNum", "Complemento", "Bairro", "Cidade", "Estado(SP)", "CEP(00000-000)"];
        List<string> cabecalhos_Recebiveis = ["CPF", "Nome", "Numero do Controle", "Recebível Exigível(R/E)", "Valor Devido", "Valor Pago", "Prazo", "Data Vencimento", "Data do Pagamento", "Emissão", "Observação Recebível", "Observação Recebido", "Tipo Pagamento", "Valor Original", "Vencimento Recebível", "Duplicata", "Parcela", "Situação", "Nome grupo", "Ordem"];
        List<string> cabecalhos_Recebidos = ["CPF", "Nome", "Numero do Controle", "Recebível Exigível(R/E)", "Valor Devido", "Valor Pago", "Prazo", "Data Vencimento", "Data do Pagamento", "Emissão", "Observação Recebido", "Tipo Pagamento", "Valor Original", "Vencimento Recebível", "Duplicata", "Parcela", "Tipo Espécie Pagamento", "Espécie Pagamento"];
        List<string> cabecalhos_Especies = ["Tipo Espécie", "Espécie Pagamento"];
        List<string> cabecalhos_Agendamentos = ["ID", "CPF", "Nome Completo", "Telefone", "Data Início (01/12/2024 00:00)", "Data Término (01/12/2024 00:00)", "Data Inclusão (01/12/2024)", "NomeCompletoDentista", "Observacao"];
        List<string> cabecalhos_Procedimentos = ["Nome Tabela", "Especialidade", "Ativo (Sim/Não)", "Nome do Procedimento", "Abreviação", "Preço", "TUSS", "Especialidade Código"];
        List<string> cabecalhos_CodProcedimentos = ["ID", "Nome", "Usuário"];
        List<string> cabecalhos_ManutencaoMan001 = ["Paciente Nome Completo", "Paciente CPF", "Observação", "Data Modificado", "Diagnostico", "Data Inicial", "Data Final"];
        List<string> cabecalhos_ManutencaoMan101 = ["Paciente CPF", "Paciente Nome", "Dentista CPF", "Dentista Nome", "Dentista Codigo", "Procedimento Nome", "Data Atendimento", "Data Início", "Data Retorno", "Procedimento Observação", "Lancamento"];
        List<string> cabecalhos_ManutencaoMan111 = ["Numero do Controle", "Paciente CPF", "Paciente Nome", "Dentista Nome", "Procedimento Nome", "Procedimento Valor", "Valor Original", "Valor do Pagamento", "Data do Pagamento", "Dente", "Procedimento Observação", "Quantidade Orto", "Tipo Pagamento", "Vencimento", "Valor Devido", "Valor Total", "Data Atendimento"];
        List<string> cabecalhos_ProcedimentosATD222 = ["Numero do Controle", "Paciente CPF", "Paciente Nome", "Dentista CPF", "Dentista Nome", "Dente", "Procedimento Nome", "Procedimento Valor", "Procedimento Observação", "Data Início", "Data Termino", "Data Atendimento"];
        List<string> cabecalhos_ManutencaoManMerge001_101 = ["Paciente Nome Man001", "Paciente CPF Man001", "Observação", "DataModificado", "Diagnostico", "DataInicial", "DataFinal", "Paciente CPF Man101", "Paciente Nome Completo", "Dentista CPF", "DentistaNome", "DentistaCodigo", "Procedimento", "DataAtendimento", "DataInicio", "DataRetorno"];

        //public static Dictionary<string, string> pessoaCSVDict;


        public Tuple<List<string[]>, List<string>> LerArquivosExcelCsv(string arquivo, Encoding encoding)
        {
            var separador = ExcelHelper.DetectarSeparadorCSV(arquivo);

            List<string> cabecalhosCSV = ExcelHelper.GetCabecalhosCSV(arquivo, separador, encoding);
            List<string[]> linhasCSV = ExcelHelper.GetLinhasCSV(arquivo, separador, cabecalhosCSV.Count(), encoding);

            return new Tuple<List<string[]>, List<string>>(linhasCSV, cabecalhosCSV);
        }

        public void RetornaProcedimentosPorTipoEntidade(List<ProcedimentosPrecosDTO> lstGruposProcedimentos, string estabelecimentoID)
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

            var pacientes = new FireBirdContext<Models.Pacientes>(_pathDB).RetornaItensBancoPorQuery(arquivoPacientesSql);

            string arquivoDDentistasSql = "Scripts\\SelectDentistas.sql";

            var dentistas = new FireBirdContext<Models.Dentistas>(_pathDB).RetornaItensBancoPorQuery(arquivoDDentistasSql);

            var lstPacientesDentistas = ConversorEntidadeParaDTO.ConvertPacientesDentistasParaPacientesDentistasDTO(pacientes, dentistas);

            var dataTablePacientesDentistas = ExcelHelper.ConversorEntidadeParaDataTable(lstPacientesDentistas);

            if (dataTablePacientesDentistas != null)
            {
                var salvarArquivoPacientesDentistas = Tools.GerarNomeArquivo($"CadastroPacienteDentistasEntidade_{Tools.ultimoEstabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoPacientesDentistas + ".xlsx", dataTablePacientesDentistas);
            }
        }              

        void IDataBaseMigracao.DataBaseImportacaoDevClinico()
        {
            ExcelHelper excelHelper = new ExcelHelper();

            string arquivoDesenvClinicoSql = "Scripts\\SelectDesenvolvimentoClinico.sql";

           var desenvClinico = new FireBirdContext<Models.DesenvolvimentoClinico>(_pathDB).RetornaItensBancoPorQuery(arquivoDesenvClinicoSql);

            string arquivoAgendamentosSql = "Scripts\\SelectAgendamentos.sql";

            var agendamentos = new FireBirdContext<Models.Agendamentos>(_pathDB).RetornaItensBancoPorQuery(arquivoAgendamentosSql);

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

            var manutencoes = new FireBirdContext<Models.Manutencoes>(_pathDB).RetornaItensBancoPorQuery(arquivoSql);

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

            RetornaProcedimentosPorTipoEntidade(lstProcedimentosPrecos, "1");
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
