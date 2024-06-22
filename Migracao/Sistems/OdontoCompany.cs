using MathNet.Numerics;
using MathNet.Numerics.Distributions;
using Migracao.Models;
using Migracao.Utils;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System.Data;
using System.Diagnostics.Metrics;
using System.Text;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.LinkLabel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace Migracao.Sistems
{
    internal class OdontoCompany
    {
        string arquivoExcelNomesUTF8 = "Files\\NomesUTF8.xlsx";

        string[] EMD101_Pacientes = ["BAIRRO", "CEP", "CGC_CPF", "CIDADE", "CLIENTE", "CELULAR", "DT_CADASTRO", "DT_NASCIMENTO", "EMAIL", "ENDERECO", "ESTADO", "FONE1", "FONE2", "FORNECEDOR", "INSC_RG", "NOME", "NUM_CONVENIO", "NUM_ENDERECO", "NUM_FICHA", "OBS1", "SEXO_M_F"];
        string[] CRD111_Recebiveis = ["AGENCIA", "AGUARDANDO_VINCULO", "ALINEA", "AXON_ID", "BANCO", "BANDA1", "BANDA2", "BANDA3", "BAIXA", "CAMPOX", "COD_CAIXA", "CODIGO_TUSS", "CONTA", "CONTA_CORRENTE", "CONTA_DOCUMENTO", "COBRADORA", "COBRANCA", "DATA_ENV_CART", "DATA_ENV_SCPC", "DATA_REMESSA", "DATA_RET_CART", "DATA_RET_SCPC", "DESCONTO_BOLETO", "DEVOLUCAO", "DOCUMENTO", "DT_AXON", "DUPLICATA", "EMISSAO", "EMITENTE", "ENCARGOS", "FILIAL", "GEROU_TRANSMISSAO", "GRUPO", "ID_BAIXAPLANOS", "ID_PIX", "JUROS", "LANCTO", "LOCAL", "LOJA", "MODIFICADO", "MOTIVO", "MULTA", "NOME_GRUPO", "NOME_LOCAL", "NOSSONUMERO", "NUM_BANCO", "OBS", "ORDEM", "PARCELA", "PERIODO", "PRAZO", "REAPRESENTOU", "RECEBEU_TRANSMISSAO", "REMESA", "RESPONSAVEL", "SEQ_ALINEA11", "SITUACAO", "SITUACAO_REMESSA", "TERMINAL", "TIPO_COBRANCA", "TIPO_DOC", "TOTAL", "TRANSMISSAO", "USUARIO", "VALOR", "VALOR_ORIG", "VALOR_RECEBER", "VALOR_VENDA", "VENCTO", "VENCTO_ORIG", "VR_CALCULADO", "VR_PARCELA"];
        string[] CXD555_Baixa = ["AGENCIA", "BANCO", "BAIXA", "CALCULO", "CNPJ_CPF", "CONTA", "DATA", "DOCUMENTO", "DT_AXON", "DT_DEPOSITO", "FECHAR_DIRETO", "FICHA_FINANCEIRO", "HISTORICO", "HORA", "LANCTO", "LOJA", "LOTE", "MODIFICADO", "NUM_CONVENIO", "OBS1", "OBS2", "OBS3", "PERIODO", "PRO_MED", "PRO_ODO", "RESPONSAVEL", "ROY_MED", "ROY_ODO", "TERMINAL", "TIPO", "TRANSMISSAO", "USUARIO", "VALOR", "VALOR_RECEBER", "VLR_BRUTO"];
        string[] BXD111_Baixa = ["AGUARDANDO_VINCULO", "AXON_ID", "BAIXA", "BANCO", "CAMPOX", "CGC_CPF", "COD_CAIXA", "CONTA_CORRENTE", "CONTA_DOCUMENTO", "DATA_REMESSA", "DOCUMENTO", "DT_AXON", "DUPLICATA", "GRUPO", "ID_BAIXAPLANOS", "LANCTO", "LOJA", "MODIFICADO", "MOTIVO", "NOME_GRUPO", "NUM_BANCO", "OBS", "PARCELA", "RESPONSAVEL", "TERMINAL", "TIPO_DOC", "TRANSMISSAO", "USUARIO", "VALOR", "VENCTO", "VR_CALCULADO", "VR_PARCELA"];
        string[] CED006_Dentistas = ["ADMISSAO", "AGENCIA", "AGENCIA2", "AXON_ID", "BAIRRO", "BANCO", "BANCO2", "CAIXA_POSTAL", "CEP", "CGC_CPF", "CIDADE", "CLIENTE", "CODIGO", "CODIGO_CLIENTE", "CODIGO_INDICACAO", "CODIGO_VALIDADE", "COD_MUNICIPIO", "COD_PRAMELHOR", "COD_UF", "COD_VENDEDOR", "CONJUGE", "CONTA", "CONTA2", "CPF_FIA", "CPF_INDICACAO", "DATA_APROVACAO_DRCASH", "DATA_BLOQUEIO", "DATA_DEP_EXCLUIDO", "DATA_LGPD", "DATA_VALIDADE", "DEPENDENTE", "DT_AXON", "DT_CADASTRO", "DT_NASC_FIA", "DT_NASCIMENTO", "DT_NASCIMENTO_DEP", "DT_ULTMOV", "EMAIL", "ENDERECO", "ENDERECO_FIA", "ESTADO", "ESTADO_FIA", "FAX", "FONE1", "FONE2", "FONE_FIA", "FONE_REF_1", "FONE_REF_2", "F_OU_J", "FORNECEDOR", "FUNCAO", "ID_DRCASH", "INSC_RG", "INSTITUTO_ODC", "LGPD_CPF", "LGPD_DATA_HORA", "LGPD_IMAGEM", "LGPD_MENSAGEM", "LGPD_TELEFONE", "LGPD_USUARIO", "LOJA", "MAE", "MODIFICADO", "NOME", "NOME_FIA", "NOME_GRUPO", "NOME_LOCAL", "NOME_REF_1", "NOME_REF_2", "NOME_VALIDADE", "NUM_BLOQUEIO", "NUM_CONVENIO", "NUM_ENDERECO", "NUM_FICHA", "OBS1", "OBS_VALIDADE", "ONDE_TRABALHA", "PAI", "PARENTESCO_FIA", "PRESTADOR", "PROFISSAO", "PROFISSAO_FIA", "PROTETICO", "PROTETICO_ATIVO", "QTDE_DEPENDENTES", "RENDA_FIA", "RENDA_MES", "RG_FIA", "SEXO_M_F", "TITULAR", "TITULAR_DEP_EXCLUIDO", "TRANSMISSAO", "USU_BLOQUEIO", "USU_CADASTRO", "USUARIO", "USUARIO_LGPD", "USUARIO_VALIDADE", "VALOR_MAXIMO_DRCASH", "VR_LIMITE"];
        string[] CED001_Procedimentos = ["CODIGO", "NOME", "SIMBOLO", "GRUPO", "CODIGO_TUSS", "VRVENDA", "ATIVO", "OBS", "PARTICULAR"];
        string[] CED002_CodProcedimentos = ["CODIGO", "NOME", "USUARIO"];
        string[] MAN001_Manutencao = ["CNPJ_CPF", "QTDE_MANUT", "OBS", "ATIVO", "INICIO_MANUT", "FINAL_MANUT", "MODIFICADO", "USUARIO", "DATA_ATIVO", "USU_ATIVO", "USU_INSTALACAO", "USU_RETIRADA", "MOTIVO_ATIVO", "DIAGNOSTICO", "PROGNOSTICO", "OBS_CLASSE", "CLASSE", "DATA_ALTERACAO", "HORA_ALTERACAO", "USUARIO_ALTERACAO", "DATA_MODIFICADO", "DT_AXON", "AXON_ID"];
        string[] MAN101_Manutencao = ["LANCTO", "NUM_MAN", "CNPJ_CPF", "CONTROLE", "DATA_PAGO", "DATA_RETORNO", "RESP_ATEND", "NOME_RESP_ATEND", "OBS_ATEND", "TIPO_ATEND", "DATA_LANC", "DATA_MANUT", "DATA_MODIFICADO"];
        string[] MAN111_Manutencao = ["LANCTO", "CNPJ_CPF", "DATA_PAGTO", "TIPO_PAGTO", "VALOR", "MES_ANO", "CAMPOX", "DIA_MES_ANO", "RESPONSAVEL", "DATA_ATEND", "RESP_ATEND", "NUM_MAN", "NOME_RESP_ATEND", "OBS", "RETORNO", "DATA_LANC", "HORA", "VENCTO", "VALOR_PARCELA", "NOSSO_NUMERO", "TIPO_MAN", "VENCTO_ORIG", "VALOR_ORIG", "MOTIVO_ALTERAR", "AUTORIZA_ALTERAR", "MOTIVO_INCLUIR", "AUTORIZA_INCLUIR", "VALOR_CALCULADO", "AUTORIZA_RECEBER", "MOTIVO_RECEBER", "EMISSAO", "BANCO", "DATA_ALTERACAO", "OPERACAO", "DOCUMENTO", "NOME_TIPO", "TIPO_COBRANCA", "TERMINAL", "REMESSA", "SENHA_ATEND", "USUARIO", "EM_ATENDIMENTO", "NOME_ATEND", "CALC_JUROS", "USU_VENDA", "COBRADORA", "COBRANCA", "DATA_REMESSA", "SITUACAO_REMESSA", "DATA_MODIFICADO", "CONTA", "NSU_TRANSACAO", "COBRANCA_ORTO", "DATA_REMESSA_CARTAO", "CONTROLE_CARTAO", "UNIDADE", "HORA_ATEND", "HORA_FINAL", "DESCONTO_BOLETO", "AGUARDANDO_VINCULO", "ID_PIX", "DT_AXON", "AXON_ID"];
        string[] ATD222_Procedimentos = ["DOCUMENTO", "CNPJ_CPF", "PRODUTO", "TIPO", "NOME_PRODUTO", "DATA", "VALOR", "OBS", "CAMPOX", "RESPONSAVEL", "DATA_ATEND", "RESP_ATEND", "NOME_RESP_ATEND", "NUMERO", "CONTROLE", "CBARRA", "TERMINAL", "USUARIO", "QTDE_SESSAO", "QTDE_SESSAO_ORIG", "MODIFICADO", "DATA_CANCELADO", "USUARIO_CANCELADO", "DATA_INCLUIDO", "USUARIO_INCLUIDO", "FORMA_MEDIDA", "QTDE_MEDIDA", "DT_AXON", "AXON_ID"];

        List<string> cabecalhos_Pacientes = ["Código", "Ativo(S/N)", "NomeCompleto", "NomeSocial", "Apelido", "Documento(CPF,CNPJ,CGC)", "DataCadastro(01/12/2024)", "Observações", "Email", "RG", "Sexo(M/F)", "NascimentoData", "NascimentoLocal", "EstadoCivil(S/C/V)", "Profissao", "CargoNaClinica", "Dentista(S/N)", "ConselhoCodigo", "Paciente(S/N)", "Funcionario(S/N)", "Fornecedor(S/N)", "TelefonePrincipal", "Celular", "TelefoneAlternativo", "Logradouro", "LogradouroNum", "Complemento", "Bairro", "Cidade", "Estado(SP)", "CEP(00000-000)"];
        List<string> cabecalhos_Recebiveis = ["CPF", "Nome", "Documento Ref", "Recebível Exigível(R/E)", "Valor Devido", "Valor Pago", "Prazo", "Vencimento", "Data do Pagamento", "Emissão", "Observação Recebível", "Observação Recebido", "Tipo do Pagamento", "Valor Original", "Vencimento Recebível", "Duplicata"];
        List<string> cabecalhos_Agendamentos = ["ID", "CPF", "Nome Completo", "Telefone", "Data Início (01/12/2024 00:00)", "Data Término (01/12/2024 00:00)", "Data Inclusão (01/12/2024)", "NomeCompletoDentista", "Observacao"];
        List<string> cabecalhos_Procedimentos = ["Nome Tabela", "Especialidade", "Ativo (Sim/Não)", "Nome do Procedimento", "Abreviação", "Preço", "TUSS", "Especialidade Código"];
        List<string> cabecalhos_CodProcedimentos = ["ID", "Nome", "Usuário"];
        List<string> cabecalhos_ManutencaoMan001 = ["Paciente Nome Completo", "Paciente CPF", "Observação", "Data Modificado", "Diagnostico", "Data Inicial", "Data Final"];
        List<string> cabecalhos_ManutencaoMan101 = ["Paciente CPF", "Paciente Nome", "Dentista CPF", "Dentista Nome", "Dentista Codigo", "Procedimento Nome", "Data Atendimento", "Data Início", "Data Retorno", "Procedimento Observação", "Lancamento"];
        List<string> cabecalhos_ManutencaoMan111 = ["Paciente CPF", "Paciente Nome", "Dentista Nome", "Procedimento Nome", "Valor Original", "Valor do Pagamento", "Data do Pagamento", "Dente", "Procedimento Observação", "Quantidade Orto", "Tipo do Pagamento", "Vencimento", "Valor Devido"];
        List<string> cabecalhos_ProcedimentosATD222 = ["Número do Controle", "Paciente CPF", "Paciente Nome", "Dentista CPF", "Dentista Nome", "Dente", "Procedimento Nome", "Procedimento Valor", "Procedimento Observação", "Data Início", "Data Termino"];
        List<string> cabecalhos_ManutencaoManMerge001_101 = ["Paciente Nome Man001", "Paciente CPF Man001", "Observação", "DataModificado", "Diagnostico", "DataInicial", "DataFinal", "Paciente CPF Man101", "Paciente Nome Completo", "Dentista CPF", "DentistaNome", "DentistaCodigo", "Procedimento", "DataAtendimento", "DataInicio", "DataRetorno"];

        //public static Dictionary<string, string> pessoaCSVDict;

        public Tuple<List<string[]>, List<string>> LerArquivosExcelCsv(string arquivo, Encoding encoding)
        {
            var separador = ExcelHelper.DetectarSeparadorCSV(arquivo);

            List<string> cabecalhosCSV = ExcelHelper.GetCabecalhosCSV(arquivo, separador, encoding);
            List<string[]> linhasCSV = ExcelHelper.GetLinhasCSV(arquivo, separador, cabecalhosCSV.Count(), encoding);

            return new Tuple<List<string[]>, List<string>>(linhasCSV, cabecalhosCSV);
        }

        public void LerArquivos(string estabelecimentoID, ListView listView = null)
        {
            ExcelHelper excelHelper = new();
            DataTable dataTablePessoas = new();
            DataTable dataTableRecebiveis = new();
            DataTable dataTableAgendamentos = new();
            DataTable dataTableProcedimentos = new();
            DataTable dataTableCodProcedimentos = new();
            DataTable dataTableManutencaoMan001 = new();
            DataTable dataTableManutencaoMan101 = new();
            DataTable dataTableManutencaoMan111 = new();
            DataTable dataTableProcedimentosATD222 = new();
            DataTable dataTableProcedimentosParticular = new();
            DataTable dataTableManutencaoMerge = new();
            DataTable dataTableProcedimentosManutencaoMerge = new();
            DataTable dataTableRecebiveisHistoricoVendas = new();

            //registroRecebivel = new HashSet<string>();

            foreach (string coluna in cabecalhos_Pacientes)
                dataTablePessoas.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_Recebiveis)
                dataTableRecebiveis.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_Recebiveis)
                dataTableRecebiveisHistoricoVendas.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_Agendamentos)
                dataTableAgendamentos.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_ManutencaoMan001)
                dataTableManutencaoMan001.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_ManutencaoMan101)
                dataTableManutencaoMan101.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_ManutencaoMan111)
                dataTableManutencaoMan111.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_ProcedimentosATD222)
                dataTableProcedimentosATD222.Columns.Add(coluna, typeof(string));

            var excel_EMD101 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("EMD101")));
            if (excel_EMD101 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_EMD101.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (EMD101_Pacientes.All(cabecalhosCSV.Contains))
                    dataTablePessoas = ConvertExcelPessoasPacientes(dataTablePessoas, cabecalhosCSV, linhasCSV);
            }

            var excel_CED006 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CED006")));
            if (excel_CED006 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_CED006.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (CED006_Dentistas.All(cabecalhosCSV.Contains))
                    dataTablePessoas = ConvertExcelPessoasDentistas(dataTablePessoas, cabecalhosCSV, linhasCSV);
            }

            if (excel_CED006 != null || excel_EMD101 != null)
            {
                var salvarArquivoPessoas = Tools.GerarNomeArquivo($"CadastroPessoas_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoPessoas + ".xlsx", dataTablePessoas);
            }



            var excel_CRD111 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CRD111")));
            if (excel_CRD111 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_CRD111.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;
                dataTableRecebiveis = ConvertExcelRecebiveis(dataTableRecebiveis, cabecalhosCSV, linhasCSV, dataTablePessoas);
            }

            var excel_BXD111 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("BXD111")));
            if (excel_BXD111 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_BXD111.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;
                dataTableRecebiveis = ConvertExcelRecebidos(dataTableRecebiveis, cabecalhosCSV, linhasCSV, dataTablePessoas);
            }

            var excel_CXD555 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CXD555")));
            if (excel_CXD555 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_CXD555.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;
                //dataTableRecebiveis = ConvertExcelRecebidos(dataTableRecebiveis, cabecalhosCSV, linhasCSV, dataTablePessoas);
                List<string[]>? lstVendaCSV = new List<string[]>();

                foreach (var linha in linhasCSV)
                {
                    if (linha[1].Contains("VENDA - "))
                        lstVendaCSV.Add(linha);
                }

                dataTableRecebiveisHistoricoVendas = ConvertExcelRecebiveisHistoricoVendas(dataTableRecebiveis, cabecalhosCSV, lstVendaCSV, dataTablePessoas);

            }

            var excel_AGENDA = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("AGENDA")));
            if (excel_AGENDA != null)
            {
                var resultado = LerArquivosExcelCsv(excel_AGENDA.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;
                dataTableAgendamentos = ConvertExcelAgendamento(dataTableAgendamentos, cabecalhosCSV, linhasCSV, dataTablePessoas);
            }

            if (excel_AGENDA != null)
            {
                var salvarArquivoAgenda = Tools.GerarNomeArquivo($"CadastroAgenda_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoAgenda + ".xlsx", dataTableAgendamentos);
            }

            var excel_CED002 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CED002")));
            if (excel_CED002 != null)
            {
                foreach (string coluna in cabecalhos_CodProcedimentos)
                    dataTableCodProcedimentos.Columns.Add(coluna, typeof(string));

                var resultado = LerArquivosExcelCsv(excel_CED002.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (CED002_CodProcedimentos.All(cabecalhosCSV.Contains))
                    dataTableCodProcedimentos = ConvertExcelGrupoProcedimentos(dataTableCodProcedimentos, cabecalhosCSV, linhasCSV);
            }

            var excel_CED001 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("CED001")));
            if (excel_CED001 != null)
            {
                foreach (string coluna in cabecalhos_Procedimentos)
                    dataTableProcedimentos.Columns.Add(coluna, typeof(string));

                var resultado = LerArquivosExcelCsv(excel_CED001.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (CED001_Procedimentos.All(cabecalhosCSV.Contains))
                    dataTableProcedimentos = ConvertExcelProcedimentos(dataTableProcedimentos, cabecalhosCSV, linhasCSV, dataTableCodProcedimentos);

                if (dataTableProcedimentos != null)
                    RetornaProcedimentosPorTipo(dataTableProcedimentos, estabelecimentoID);
            }

            var excel_MAN001 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("MAN001")));

            if (excel_MAN001 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_MAN001.Text, Encoding.UTF8);
                var linhasCSV1 = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (MAN001_Manutencao.All(cabecalhosCSV.Contains))
                    dataTableManutencaoMan001 = ConvertExcelManutencaoMAM001(dataTableManutencaoMan001, cabecalhosCSV, linhasCSV1, dataTablePessoas);

                if (dataTableManutencaoMan001 != null)
                {
                    var salvarManutencaoMan001 = Tools.GerarNomeArquivo($"CadastroManutencaoMan001_{estabelecimentoID}_OdontoCompany");
                    excelHelper.CriarExcelArquivo(salvarManutencaoMan001 + ".xlsx", dataTableManutencaoMan001);
                }
            }

            var excel_MAN101 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("MAN101")));

            if (excel_MAN101 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_MAN101.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (MAN101_Manutencao.All(cabecalhosCSV.Contains))
                    dataTableManutencaoMan101 = ConvertExcelProcedimentosMAN101(dataTableManutencaoMan101, cabecalhosCSV, linhasCSV, dataTablePessoas);

                if (dataTableManutencaoMan101 != null)
                {
                    var salvarManutencaoMan101 = Tools.GerarNomeArquivo($"CadastroManutencaoMan101_{estabelecimentoID}_OdontoCompany");
                    excelHelper.CriarExcelArquivo(salvarManutencaoMan101 + ".xlsx", dataTableManutencaoMan101);
                }

                //if (dataTableManutencaoMan001 != null && dataTableManutencaoMan101 != null)
                //    dataTableManutencaoMerge = MergeDataTables(dataTableManutencaoMan001, dataTableManutencaoMan101);


                //if (dataTableManutencaoMerge != null)
                //{
                //    var salvarManutencaoMergeMan = Tools.GerarNomeArquivo($"CadastroManutencaoManMerge_001_101_{estabelecimentoID}_OdontoCompany");
                //    excelHelper.CriarExcelArquivo(salvarManutencaoMergeMan + ".xlsx", dataTableManutencaoMerge);
                //}
            }

            var excel_ATD222 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("ATD222")));

            if (excel_ATD222 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_ATD222.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (ATD222_Procedimentos.All(cabecalhosCSV.Contains))
                    dataTableProcedimentosATD222 = ConvertExcelProcedimentosATD222(dataTableProcedimentosATD222, cabecalhosCSV, linhasCSV, dataTablePessoas, dataTableRecebiveis);

                if (dataTableProcedimentosATD222 != null)
                {
                    var salvarProcedimentosATD222 = Tools.GerarNomeArquivo($"CadastroProcedimentosATD222_{estabelecimentoID}_OdontoCompany");
                    excelHelper.CriarExcelArquivo(salvarProcedimentosATD222 + ".xlsx", dataTableProcedimentosATD222);
                }

            }

            var excel_MAN111 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("MAN111")));

            if (excel_MAN111 != null)
            {
                var resultado = LerArquivosExcelCsv(excel_MAN111.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (MAN111_Manutencao.All(cabecalhosCSV.Contains))
                    dataTableManutencaoMan111 = ConvertExcelProcedimentosMAN111(dataTableManutencaoMan111, cabecalhosCSV, linhasCSV, dataTablePessoas, dataTableManutencaoMan101);

                if (dataTableManutencaoMan111 != null)
                {
                    var salvarManutencaoMan111 = Tools.GerarNomeArquivo($"CadastroManutencaoMan111_{estabelecimentoID}_OdontoCompany");
                    excelHelper.CriarExcelArquivo(salvarManutencaoMan111 + ".xlsx", dataTableManutencaoMan111);
                }

                if (dataTableManutencaoMan111 != null && dataTableProcedimentosATD222 != null)
                    dataTableProcedimentosManutencaoMerge = MergeDataTablesProcedimentosManutencao(dataTableProcedimentosATD222, dataTableManutencaoMan111);


                if (dataTableProcedimentosManutencaoMerge != null)
                {
                    var salvarProcedimentosManutencaoMerge = Tools.GerarNomeArquivo($"cadastroProcedimentosManutencaoMerge_{estabelecimentoID}_odontocompany");
                    excelHelper.CriarExcelArquivo(salvarProcedimentosManutencaoMerge + ".xlsx", dataTableProcedimentosManutencaoMerge);
                }
            }

            if (dataTableProcedimentos.Rows.Count > 0)
            {
                var salvarArquivoAgenda = Tools.GerarNomeArquivo($"CadastroProcedimentos_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoAgenda + ".xlsx", dataTableProcedimentos);
            }

            if (excel_BXD111 != null || excel_CRD111 != null)
            {
                var salvarArquivoRecebiveis = Tools.GerarNomeArquivo($"CadastroRecebiveis_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoRecebiveis + ".xlsx", dataTableRecebiveis);
            }

            if (excel_CXD555 != null)
            {
                var salvarArquivoRecebiveisHistoricoVendas = Tools.GerarNomeArquivo($"CadastroRecebiveisHistoricoVendas_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoRecebiveisHistoricoVendas + ".xlsx", dataTableRecebiveisHistoricoVendas);
            }
        }

        public DataTable ConvertExcelRecebiveis(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePacientes)
        {
            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CGC_CPF").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();
                        var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
                        var vencimentoData = valoresLinha.GetValueOrDefault("VENCTO").Trim();
                        var emissaoData = valoresLinha.GetValueOrDefault("EMISSAO").Trim();
                        var tipoDoc = valoresLinha.GetValueOrDefault("TIPO_DOC").Trim();
                        var valorOriginal = valoresLinha.GetValueOrDefault("VALOR_ORIG").Trim();
                        var venctoOriginal = valoresLinha.GetValueOrDefault("VENCTO_ORIG").Trim();
                        var duplicata = valoresLinha.GetValueOrDefault("DUPLICATA").Trim();

                        cpf = cpf.ToCPF();
                        string nome = "";

                        if (dataTablePacientes.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePacientes.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();
                        }

                        if (!string.IsNullOrEmpty(valor) && valor.ArredondarValorV2() > 1)
                        {
                            dataRow["CPF"] = cpf;
                            dataRow["Nome"] = nome;
                            dataRow["Observação Recebível"] = observacao;
                            dataRow["Documento Ref"] = documento;
                            dataRow["Valor Devido"] = valor.ArredondarValorV2();
                            //dataRow["Prazo"] = prazo;
                            dataRow["Vencimento"] = vencimentoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["Emissão"] = emissaoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["Recebível Exigível(R/E)"] = "R";
                            dataRow["Duplicata"] = duplicata;
                            dataRow["Tipo do Pagamento"] = tipoDoc;
                            dataRow["Valor Original"] = valorOriginal;
                            dataRow["Vencimento Recebível"] = venctoOriginal.ToData().ToString("dd/MM/yyyy");

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
            }
        }
        public DataTable ConvertExcelRecebiveisHistoricoVendas(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePacientes)
        {
            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("HISTORICO").Trim();
                        var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();
                        var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
                        var vencimentoData = valoresLinha.GetValueOrDefault("DATA").Trim();
                        var emissaoData = valoresLinha.GetValueOrDefault("TRANSMISSAO").Trim();

                        cpf = cpf.ToCPF();
                        var nome = observacao.Replace("VENDA - ", "").Trim();

                        if (dataTablePacientes.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePacientes.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();
                        }

                        if (!string.IsNullOrEmpty(valor) && valor.ArredondarValorV2() > 1)
                        {
                            dataRow["CPF"] = cpf;
                            dataRow["Nome"] = nome;
                            dataRow["Observação Recebível"] = observacao;
                            dataRow["Documento Ref"] = documento;
                            dataRow["Valor Original"] = valor.ArredondarValorV2();
                            //dataRow["Prazo"] = prazo;
                            dataRow["Vencimento"] = vencimentoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["Emissão"] = emissaoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["Recebível Exigível(R/E)"] = "R";

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
            }
        }
        public DataTable ConvertExcelRecebidos(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePacientes)
        {
            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();  // Vem do CRD111
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++) // Vem do BXD111
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cabecalhoCpf = "CGC_CPF";
                        var cabecalhoObs = "OBS";
                        var cabecalhoDataBaixa = "BAIXA";
                        var cabecalhoDataVencimento = "VENCTO";
                        var cabecalhoParcela = "VR_PARCELA";

                        if (cabecalhos.Contains("CNPJ_CPF"))
                        {
                            cabecalhoCpf = "CNPJ_CPF";
                            cabecalhoObs = "OBS1";
                            cabecalhoDataBaixa = "DATA";
                            cabecalhoDataVencimento = "DATA";
                            cabecalhoParcela = "VALOR";
                        }

                        var cpf = valoresLinha.GetValueOrDefault(cabecalhoCpf).Trim();
                        var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();
                        var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
                        var valorOriginal = valoresLinha.GetValueOrDefault(cabecalhoParcela).Trim();
                        var observacao = valoresLinha.GetValueOrDefault(cabecalhoObs).Trim();
                        var baixaData = valoresLinha.GetValueOrDefault(cabecalhoDataBaixa).Trim();
                        var vencimentoData = valoresLinha.GetValueOrDefault(cabecalhoDataVencimento).Trim();

                        if (!string.IsNullOrEmpty(valor) && !string.IsNullOrEmpty(valorOriginal) && valor.ArredondarValorV2() > 1 && valorOriginal.ArredondarValorV2() > 1)
                        {
                            cpf = cpf.ToCPF();
                            string nome = "";

                            if (dataTablePacientes.Rows.Count > 0)
                            {
                                DataRow[] dataRowPessoasEncontrados = dataTablePacientes.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                                if (dataRowPessoasEncontrados.Length > 0)
                                    nome = dataRowPessoasEncontrados[0]["NomeCompleto"].ToString();
                            }

                            DataRow[] dataRowEncontrados = dataTable.AsEnumerable()
                            .Where(row =>
                                row.Field<string>("Documento Ref") == documento &&
                                row.Field<string>("CPF") == cpf &&
                                row.Field<string>("Valor Devido") == valorOriginal.ArredondarValorV2().ToString())

                            .ToArray();

                            dataRow["CPF"] = cpf;
                            dataRow["Nome"] = nome;
                            dataRow["Documento Ref"] = documento;
                            dataRow["Valor Pago"] = valor.ArredondarValorV2();
                            dataRow["Data do Pagamento"] = baixaData;
                            dataRow["Vencimento"] = vencimentoData.ToData();
                            dataRow["Observação Recebido"] = observacao;
                            dataRow["Recebível Exigível(R/E)"] = "R";

                            if (dataRowEncontrados.Length > 0)
                            {
                                dataRow["Valor Devido"] = dataRowEncontrados[0]["Valor Devido"].ToString();
                                dataRow["Tipo do Pagamento"] = dataRowEncontrados[0]["Tipo do Pagamento"].ToString();
                                dataRow["Valor Original"] = dataRowEncontrados[0]["Valor Original"].ToString();
                                dataRow["Vencimento Recebível"] = dataRowEncontrados[0]["Vencimento Recebível"].ToString();
                                dataRow["Duplicata"] = dataRowEncontrados[0]["Duplicata"].ToString();
                            }                               

                            dataTable.Rows.Add(dataRow);

                            //if (dataRowEncontrados.Length > 0)
                            //{
                            //    foreach (var dataRowEncontrado in dataRowEncontrados)
                            //    {
                            //        dataRowEncontrado["Valor Pago"] = valor.ArredondarValorV2();
                            //        dataRowEncontrado["Data do Pagamento"] = baixaData.ToData().ToString("dd/MM/yyyy");
                            //        dataRowEncontrado["Observação Recebido"] = observacao;
                            //    }
                            //}

                            //else
                            //{
                            //    dataRow["CPF"] = cpf;
                            //    dataRow["Nome"] = nome;
                            //    dataRow["Documento Ref"] = documento;
                            //    dataRow["Valor Pago"] = valor.ArredondarValorV2();
                            //    dataRow["Data do Pagamento"] = baixaData;
                            //    dataRow["Vencimento"] = vencimentoData.ToData();
                            //    dataRow["Observação Recebido"] = observacao;
                            //    dataRow["Recebível Exigível(R/E)"] = "R";

                            //    dataTable.Rows.Add(dataRow);
                            //}
                        }
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebidos: {error.Message}");
            }
        }
        public DataTable ConvertExcelPessoasDentistas(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas)
        {
            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length) // Verificar se o índice está dentro do tamanho da linha
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        //if (fichasCadastradas.Contains(codigo.ToNum()))
                        var codigo = valoresLinha.GetValueOrDefault("CODIGO").Trim();
                        var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
                        var departamento = valoresLinha.GetValueOrDefault("DEPARTAMENTO").Trim();
                        var obs = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var ativo = valoresLinha.GetValueOrDefault("ATIVO").Trim();
                        var nomeCompleto = valoresLinha.GetValueOrDefault("NOME_COMPLETO").Trim();
                        var email = valoresLinha.GetValueOrDefault("EMAIL").Trim();
                        var telefone = valoresLinha.GetValueOrDefault("TELEFONE").Trim();
                        var cro = valoresLinha.GetValueOrDefault("CRO").Trim();
                        var modificado = valoresLinha.GetValueOrDefault("MODIFICADO").Trim();

                        //pessoaCSVDict.Add("dentista|" + nome, nome.ToNome());

                        dataRow["Código"] = codigo.ToNum();
                        dataRow["Ativo(S/N)"] = "S";
                        dataRow["NomeCompleto"] = nome.ToNome();
                        dataRow["NomeSocial"] = "";
                        dataRow["Apelido"] = nome.GetPrimeirosCaracteres(20).ToNome();
                        dataRow["DataCadastro(01/12/2024)"] = modificado.ToData();
                        dataRow["Observações"] = obs;
                        dataRow["Email"] = email.ToEmail();
                        dataRow["NascimentoLocal"] = "";
                        dataRow["EstadoCivil(S/C/V)"] = "";
                        dataRow["Profissao"] = "";
                        dataRow["CargoNaClinica"] = "";
                        dataRow["Dentista(S/N)"] = "N";
                        dataRow["ConselhoCodigo"] = "";
                        dataRow["Paciente(S/N)"] = "N";
                        dataRow["Funcionario(S/N)"] = "S";
                        dataRow["Fornecedor(S/N)"] = "N";

                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Dentistas: {error.Message}");
            }
        }
        public DataTable ConvertExcelPessoasPacientes(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas)
        {
            ExcelHelper excelHelper = new();
            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length) // Verificar se o índice está dentro do tamanho da linha
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var numFicha = valoresLinha.GetValueOrDefault("NUM_FICHA").Trim();
                        var cliente = valoresLinha.GetValueOrDefault("CLIENTE").Trim();
                        var fornecedor = valoresLinha.GetValueOrDefault("FORNECEDOR").Trim();
                        var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
                        var cgcCpf = valoresLinha.GetValueOrDefault("CGC_CPF").Trim();
                        var rg = valoresLinha.GetValueOrDefault("INSC_RG").Trim();
                        var sexo = valoresLinha.GetValueOrDefault("SEXO_M_F").Trim();
                        var email = valoresLinha.GetValueOrDefault("EMAIL").Trim();
                        var fone1 = valoresLinha.GetValueOrDefault("FONE1").Trim();
                        var fone2 = valoresLinha.GetValueOrDefault("FONE2").Trim();
                        var celular = valoresLinha.GetValueOrDefault("CELULAR").Trim();
                        var endereco = valoresLinha.GetValueOrDefault("ENDERECO").Trim();
                        var bairro = valoresLinha.GetValueOrDefault("BAIRRO").Trim();
                        var numEndereco = valoresLinha.GetValueOrDefault("NUM_ENDERECO").Trim();
                        var cidade = valoresLinha.GetValueOrDefault("CIDADE").Trim();
                        var estado = valoresLinha.GetValueOrDefault("ESTADO").Trim();
                        var cep = valoresLinha.GetValueOrDefault("CEP").Trim();
                        var obs = valoresLinha.GetValueOrDefault("OBS1").Trim();
                        var numConvenio = valoresLinha.GetValueOrDefault("NUM_CONVENIO").Trim();
                        var dataCadastro = valoresLinha.GetValueOrDefault("DT_CADASTRO").Trim();
                        var dataNascimento = valoresLinha.GetValueOrDefault("DT_NASCIMENTO").Trim();

                        //if (cliente == "S")
                        //	pessoas.Add("paciente|" + nome, numFicha);

                        //if (cliente != "S" && fornecedor != "S")
                        cliente = "S";

                        if (!excelHelper.CidadeExists(cidade.PrimeiraLetraMaiuscula(), estado))
                            cidade = cidade.EncontrarCidadeSemelhante();

                        //dataRow["Codigo"] = numFicha.ToNum();
                        dataRow["Ativo(S/N)"] = "S";
                        dataRow["NomeCompleto"] = nome.ToNome();
                        dataRow["NomeSocial"] = "";
                        dataRow["Apelido"] = nome.GetPrimeirosCaracteres(20).ToNome();
                        dataRow["Documento(CPF,CNPJ,CGC)"] = cgcCpf.ToCPF();
                        dataRow["DataCadastro(01/12/2024)"] = dataCadastro.ToData();
                        dataRow["Observações"] = obs;
                        dataRow["Email"] = email.ToEmail();
                        dataRow["RG"] = rg.GetPrimeirosCaracteres(20);
                        dataRow["Sexo(M/F)"] = sexo.ToSexo("m", "f") ? "M" : "F";
                        dataRow["NascimentoData"] = dataNascimento.ToDataNull();
                        dataRow["NascimentoLocal"] = "";
                        dataRow["EstadoCivil(S/C/V)"] = "";
                        dataRow["Profissao"] = "";
                        dataRow["CargoNaClinica"] = "";
                        dataRow["Dentista(S/N)"] = "";
                        dataRow["ConselhoCodigo"] = "";
                        dataRow["Paciente(S/N)"] = cliente;
                        dataRow["Funcionario(S/N)"] = "N";
                        dataRow["Fornecedor(S/N)"] = fornecedor;
                        dataRow["TelefonePrincipal"] = fone1.ToFone();
                        dataRow["Celular"] = celular.ToFone();
                        dataRow["TelefoneAlternativo"] = fone2.ToFone();
                        dataRow["Logradouro"] = endereco.PrimeiraLetraMaiuscula();
                        dataRow["LogradouroNum"] = numEndereco;
                        dataRow["Complemento"] = "";
                        dataRow["Bairro"] = bairro.PrimeiraLetraMaiuscula();
                        dataRow["Cidade"] = cidade.EncontrarCidadeSemelhante();
                        dataRow["Estado(SP)"] = estado.ToUpper();
                        dataRow["CEP(00000-000)"] = cep.ToNum();

                        dataTable.Rows.Add(dataRow);
                    }

                    catch (Exception error)
                    {
                        //throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
            }
        }
        public DataTable ConvertExcelGrupoProcedimentos(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas)
        {
            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cod = valoresLinha.GetValueOrDefault("CODIGO").Trim();
                        var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
                        var usuario = valoresLinha.GetValueOrDefault("USUARIO").Trim();

                        dataRow["ID"] = cod;
                        dataRow["Nome"] = nome.GetPrimeirosCaracteres(100).PrimeiraLetraMaiuscula();
                        dataRow["Usuário"] = nome.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();

                        dataTable.Rows.Add(dataRow);
                    }

                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Grupo Procedimentos: {error.Message}");
            }
        }
        public DataTable ConvertExcelProcedimentos(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable codProcedimentos)
        {
            try
            {
                var grupoCategoriaDict = new Dictionary<string, string[]>()
                {
                    { "CIRURGIA", new string[] { "Cirurgia", ((byte)ProcedimentosCategorias.Cirurgia).ToString(), ((byte)ProcedimentosCategoriasID.Cirurgia).ToString() } },
                    { "ENDODONTIA", new string[] { "Endodontia", ((byte)ProcedimentosCategorias.Endodontia).ToString(), ((byte)ProcedimentosCategoriasID.Endodontia).ToString() } },
                    { "PERIODONTIA", new string[] { "Periodontia", ((byte)ProcedimentosCategorias.Periodontia).ToString(), ((byte)ProcedimentosCategoriasID.Periodontia).ToString() } },
                    { "PROTESE", new string[] { "Protese", ((byte)ProcedimentosCategorias.Prótese).ToString(), ((byte)ProcedimentosCategoriasID.Prótese).ToString() } },
                    { "CLINICO", new string[] { "Outros", ((byte)ProcedimentosCategorias.Outros).ToString(), ((byte)ProcedimentosCategoriasID.Outros).ToString() } },
                    { "MANUTENCAO", new string[] { "Ortodontia", ((byte)ProcedimentosCategorias.Ortodontia).ToString(), ((byte)ProcedimentosCategoriasID.Ortodontia).ToString() } },
                    { "ORTODONTIA", new string[] { "Ortodontia", ((byte)ProcedimentosCategorias.Ortodontia).ToString(), ((byte)ProcedimentosCategoriasID.Ortodontia).ToString() } },
                    { "ORTO", new string[] { "Ortodontia", ((byte)ProcedimentosCategorias.Ortodontia).ToString(), ((byte)ProcedimentosCategoriasID.Ortodontia).ToString() } },
                    { "PREVENCAO", new string[] { "Prevencao", ((byte)ProcedimentosCategorias.Prevenção).ToString(), ((byte)ProcedimentosCategoriasID.Prevenção).ToString() } },
                    { "OROFACIAL", new string[] { "Orofacial", ((byte)ProcedimentosCategorias.Orofacial).ToString(), ((int)ProcedimentosCategoriasID.Orofacial).ToString() } },
                    { "HARMONIZACAO OROFACIAL", new string[] { "Orofacial", ((byte)ProcedimentosCategorias.Outros).ToString(), ((byte)ProcedimentosCategoriasID.Outros).ToString() } }
                };
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cod = valoresLinha.GetValueOrDefault("GRUPO").Trim();
                        var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
                        var abreviacao = valoresLinha.GetValueOrDefault("SIMBOLO").Trim();
                        var tuss = valoresLinha.GetValueOrDefault("CODIGO_TUSS").Trim();
                        var valor = valoresLinha.GetValueOrDefault("VRVENDA").Trim();
                        var ativo = valoresLinha.GetValueOrDefault("ATIVO").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var particular = valoresLinha.GetValueOrDefault("PARTICULAR").Trim();
                        var nomeTabela = "";
                        var especialidade = "";
                        var especialidadeCod = "";
                        var especialidadeCodID = "";

                        DataRow[] dataRowEncontrados = codProcedimentos.AsEnumerable()
                        .Where(row =>
                            row.Field<string>("ID") == cod)
                        .ToArray();

                        //DataRow[] dataRowEncontrados = codProcedimentos.Select($"ID = '{cod}'");
                        if (dataRowEncontrados.Length > 0)
                            especialidade = dataRowEncontrados[0]["Nome"].ToString();

                        if (nome.StartsWith("odc ", StringComparison.CurrentCultureIgnoreCase))
                            nomeTabela = "ODC";
                        //else if (string.IsNullOrEmpty(observacao) && particular != "N")
                        else if (string.IsNullOrEmpty(observacao) || grupoCategoriaDict.ContainsKey(Tools.RemoverAcentos(especialidade).ToUpper()))
                            nomeTabela = "Particular";
                        else
                        {
                            nomeTabela = especialidade;
                            especialidade = "Outros";
                        }

                        var buscarEspecialidade = Tools.RemoverAcentos(especialidade).ToUpper();

                        if (grupoCategoriaDict.ContainsKey(buscarEspecialidade))
                        {
                            var grupoCategoriaEncontrado = grupoCategoriaDict[buscarEspecialidade];
                            especialidade = grupoCategoriaEncontrado[0];
                            especialidadeCod = grupoCategoriaEncontrado[1];
                            especialidadeCodID = grupoCategoriaEncontrado[2];
                        }

                        if (nomeTabela.GetPrimeirosCaracteres(40).PrimeiraLetraMaiuscula() == "")
                            nomeTabela = nomeTabela;

                        List<string> cabecalhos_Procedimentos = ["Nome Tabela", "Especialidade", "Ativo (Sim/Não)", "Nome do Procedimento", "Abreviação", "Preço", "TUSS", "Especialidade Código"];


                        dataRow["Nome Tabela"] = nomeTabela.GetPrimeirosCaracteres(40).PrimeiraLetraMaiuscula();
                        dataRow["Especialidade"] = especialidade;
                        dataRow["Ativo (Sim/Não)"] = ativo == "N" ? "N" : "S";
                        dataRow["Nome do Procedimento"] = nome.GetPrimeirosCaracteres(100).PrimeiraLetraMaiuscula();
                        dataRow["Abreviação"] = abreviacao.GetPrimeirosCaracteres(10);
                        dataRow["Preço"] = valor.ArredondarValorV2();
                        dataRow["TUSS"] = tuss.ToNumV2();
                        dataRow["Especialidade Código"] = especialidadeCodID;

                        dataTable.Rows.Add(dataRow);
                    }

                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Procedimentos: {error.Message}");
            }
        }
        public DataTable ConvertExcelAgendamento(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
        {
            List<string> ids = new();

            try
            {
                int linhaIndex = 0;
                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length) // Verificar se o índice está dentro do tamanho da linha
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var nome = valoresLinha.GetValueOrDefault("NOME").Trim();
                        var data = valoresLinha.GetValueOrDefault("DATA").Trim();
                        var hora = valoresLinha.GetValueOrDefault("HORA").Trim();
                        var cod_responsavel = valoresLinha.GetValueOrDefault("CODIGO_RESP").Trim();
                        var responsavel = valoresLinha.GetValueOrDefault("RESPONSAVEL").Trim();
                        var telefone = valoresLinha.GetValueOrDefault("FONE_1").Trim();
                        var dataInclusao = valoresLinha.GetValueOrDefault("MODIFICADO").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var id = valoresLinha.GetValueOrDefault("LANCTO").Trim();

                        if (!ids.Contains(id))
                        {
                            ids.Add(id);

                            //DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable()
                            //	.Where(row =>
                            //	row.Field<string>("Código") == cod_responsavel)
                            //	.ToArray();

                            //if (dataRowEncontrados.Length > 0)
                            //	responsavel = dataRowEncontrados[0]["NomeCompleto"].ToString();

                            var minutos = hora.Split(':')[1];
                            var horas = hora.Split(':')[0];
                            var dataInicio = data.ToData();

                            if (!string.IsNullOrEmpty(horas))
                                dataInicio = dataInicio.AddHours(double.Parse(horas));
                            if (!string.IsNullOrEmpty(minutos))
                                dataInicio = dataInicio.AddMinutes(double.Parse(minutos));

                            var dataTermino = dataInicio;
                            var idsEncontrados = linhas.Where(linha => linha[0].Equals(id)).Count();
                            if (idsEncontrados > 0)
                                dataTermino = dataTermino.AddMinutes(15 * idsEncontrados);
                            else
                                dataTermino = dataTermino.AddMinutes(15);

                            dataRow["ID"] = id;
                            dataRow["CPF"] = cpf.ToCPF();
                            dataRow["Nome Completo"] = nome.ToNome();
                            dataRow["Data Início (01/12/2024 00:00)"] = dataInicio;
                            dataRow["Data Término (01/12/2024 00:00)"] = dataTermino;
                            dataRow["Data Inclusão (01/12/2024)"] = dataInclusao.ToData();
                            dataRow["NomeCompletoDentista"] = responsavel;
                            dataRow["Telefone"] = telefone.ToFone();
                            dataRow["Observacao"] = observacao;

                            dataTable.Rows.Add(dataRow);
                        }
                    }

                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel para Pessoas Pacientes: {error.Message}");
            }
        }
        public DataTable ConvertExcelDesenvolvimentoClinicoMan001(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
        {
            try
            {
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var diagnostico = valoresLinha.GetValueOrDefault("DIAGNOSTICO").Trim();
                        var obsClasse = valoresLinha.GetValueOrDefault("OBS_CLASSE").Trim();
                        var dataModificado = valoresLinha.GetValueOrDefault("DATA_MODIFICADO").Trim();
                        var dataInicial = valoresLinha.GetValueOrDefault("INICIO_MANUT").Trim();
                        var dataFinal = valoresLinha.GetValueOrDefault("FINAL_MANUT").Trim();


                        cpf = cpf.ToCPF();
                        string? nome = null;

                        if (dataTablePessoas.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();
                        }

                        dataRow["Nome"] = nome;
                        dataRow["CNPJ_CPF"] = cpf;
                        dataRow["Observação"] = observacao + " - " + diagnostico + " - " + obsClasse; ;
                        dataRow["Diagnostico"] = diagnostico;
                        dataRow["DataModificado"] = dataModificado.ToData().ToString();
                        dataRow["DataInicial"] = dataInicial.ToData().ToString();
                        dataRow["DataFinal"] = dataFinal.ToData().ToString();

                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
            }
        }
        public DataTable ConvertExcelDesenvolvimentoClinicoMan101(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
        {
            try
            {
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var dataRetorno = valoresLinha.GetValueOrDefault("DATA_RETORNO").Trim();
                        var respAtendimento = valoresLinha.GetValueOrDefault("RESP_ATEND").Trim();
                        var nomeRespAtendimento = valoresLinha.GetValueOrDefault("NOME_RESP_ATEND").Trim();
                        var tipoAtendimento = valoresLinha.GetValueOrDefault("TIPO_ATEND").Trim();
                        var dataLancamento = valoresLinha.GetValueOrDefault("DATA_LANC").Trim();


                        dataRow["DataRetorno"] = dataRetorno.ToData().ToString();
                        dataRow["RespAtendimento"] = respAtendimento;
                        dataRow["NomeRespAtendimento"] = nomeRespAtendimento;
                        dataRow["TipoAtendimento"] = tipoAtendimento;
                        dataRow["DataLancamento"] = dataLancamento;

                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
            }
        }
        public DataTable ConvertExcelManutencaoMAM001(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
        {
            try
            {
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var diagnostico = valoresLinha.GetValueOrDefault("DIAGNOSTICO").Trim();
                        var obsClasse = valoresLinha.GetValueOrDefault("OBS_CLASSE").Trim();
                        var dataModificado = valoresLinha.GetValueOrDefault("DATA_MODIFICADO").Trim();
                        var dataInicial = valoresLinha.GetValueOrDefault("INICIO_MANUT").Trim();
                        var dataFinal = valoresLinha.GetValueOrDefault("FINAL_MANUT").Trim();


                        cpf = cpf.ToCPF();
                        string? nome = null;

                        if (dataTablePessoas.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();
                        }


                        dataRow["Paciente CPF"] = cpf;
                        dataRow["Paciente Nome Completo"] = nome;
                        dataRow["Diagnostico"] = diagnostico;
                        dataRow["Observação"] = observacao + " - " + diagnostico + " - " + obsClasse;
                        dataRow["Data Modificado"] = dataModificado.ToData().ToString();
                        dataRow["Data Inicial"] = dataInicial.ToData().ToString();
                        dataRow["Data Final"] = dataFinal.ToData().ToString();

                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
            }
        }
        public DataTable ConvertExcelProcedimentosMAN101(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
        {
            try
            {
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var dataRetorno = valoresLinha.GetValueOrDefault("DATA_RETORNO").Trim();
                        var respAtendimento = valoresLinha.GetValueOrDefault("RESP_ATEND").Trim();
                        var nomeRespAtendimento = valoresLinha.GetValueOrDefault("NOME_RESP_ATEND").Trim();
                        var procedimento = valoresLinha.GetValueOrDefault("TIPO_ATEND").Trim();
                        var dataModificado = valoresLinha.GetValueOrDefault("DATA_MODIFICADO").Trim();
                        var dataManutencao = valoresLinha.GetValueOrDefault("DATA_MANUT").Trim();
                        var obsAtendimento = valoresLinha.GetValueOrDefault("OBS_ATEND").Trim();
                        var lancamento = valoresLinha.GetValueOrDefault("LANCTO").Trim();

                        cpf = cpf.ToCPF();
                        string? nome = null;

                        if (dataTablePessoas.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();
                        }

                        dataRow["Paciente CPF"] = cpf;
                        dataRow["Paciente Nome"] = nome;
                        dataRow["Dentista CPF"] = string.Empty;
                        dataRow["Dentista Nome"] = nomeRespAtendimento;
                        dataRow["Dentista Codigo"] = respAtendimento;
                        dataRow["Procedimento Nome"] = procedimento;
                        dataRow["Data Atendimento"] = dataModificado.ToData().ToShortDateString();
                        dataRow["Data Início"] = dataManutencao.ToData().ToShortDateString();
                        dataRow["Data Retorno"] = dataRetorno.ToData().ToShortDateString();
                        dataRow["Procedimento Observação"] = obsAtendimento;
                        dataRow["Lancamento"] = lancamento;

                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel AtendimentoATD222: {error.Message}");
            }
        }
        public DataTable ConvertExcelProcedimentosMAN111(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null, DataTable dataTableMan101 = null)
        {
            try
            {
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var valorOriginal = valoresLinha.GetValueOrDefault("VALOR_ORIG").Trim();
                        var nomeRespAtendimento = valoresLinha.GetValueOrDefault("NOME_RESP_ATEND").Trim();
                        var dataPagamento = valoresLinha.GetValueOrDefault("DATA_PAGTO").Trim();
                        var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();
                        var nomeTipo = valoresLinha.GetValueOrDefault("NOME_TIPO").Trim();
                        var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
                        var lancamento = valoresLinha.GetValueOrDefault("LANCTO").Trim();
                        var tipoPagamento = valoresLinha.GetValueOrDefault("TIPO_PAGTO").Trim();
                        var venctoOriginal = valoresLinha.GetValueOrDefault("VENCTO_ORIG").Trim();
                        var valorParcela = valoresLinha.GetValueOrDefault("VALOR_PARCELA").Trim();

                        /*    
                            TIPO_PAGTO=Tipo do Pagamento 
                            VALOR=Valor do Pagamento 
                            VALOR_ORIG = Valor Original
                            VENCTO_ORIG = Vencimento
                            VALOR_PARCELA = Valor Devido
                         */

                        cpf = cpf.ToCPF();
                        string? nome = null;
                        string? procObservacao = null;


                        var docsEncontrados = linhas.Where(linha => linha[34].Equals(documento)).Count();

                        if (dataTablePessoas.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();

                            //`Para analisar porque está vazio
                            if (!string.IsNullOrEmpty(nome))
                                continue;
                        }

                        if (dataTableMan101.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTableMan101.AsEnumerable()
                                .Where(row => row.Field<string>("Paciente CPF") == cpf
                                 && row.Field<string>("Lancamento") == lancamento)
                                .ToArray();
                            if (dataRowEncontrados.Length > 0)
                            {
                                procObservacao = dataRowEncontrados[0]["Procedimento Observação"].ToString();// [9]
                            }

                        }

                        dataRow["Paciente CPF"] = cpf;
                        dataRow["Paciente Nome"] = nome;
                        dataRow["Dentista Nome"] = nomeRespAtendimento;
                        dataRow["Procedimento Nome"] = nomeTipo;
                        dataRow["Valor Original"] = valorOriginal;
                        dataRow["Valor do Pagamento"] = valor.ArredondarValorV2();
                        dataRow["Data do Pagamento"] = dataPagamento.ToDataNull()?.ToShortDateString();
                        dataRow["Dente"] = string.Empty;
                        dataRow["Procedimento Observação"] = procObservacao;
                        dataRow["Quantidade Orto"] = docsEncontrados;
                        dataRow["Tipo do Pagamento"] = tipoPagamento;
                        dataRow["Vencimento"] = venctoOriginal;
                        dataRow["Valor Devido"] = valorParcela;

                        dataTable.Rows.Add(dataRow);
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel AtendimentoATD222: {error.Message}");
            }
        }
        public DataTable ConvertExcelProcedimentosATD222(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null, DataTable dataTableRecebiveis = null)
        {
            try
            {
                int linhaIndex = 0;

                foreach (string[] linha in linhas)
                {
                    try
                    {
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
                            if (i < linha.Length)
                                valoresLinha.Add(cabecalhos[i], linha[i]);

                        var cpf = valoresLinha.GetValueOrDefault("CNPJ_CPF").Trim();
                        var nomeProduto = valoresLinha.GetValueOrDefault("NOME_PRODUTO").Trim();
                        var valor = valoresLinha.GetValueOrDefault("VALOR").Trim();
                        var observacao = valoresLinha.GetValueOrDefault("OBS").Trim();
                        var dataAtendimento = valoresLinha.GetValueOrDefault("DATA_ATEND").Trim();
                        var dataInicio = valoresLinha.GetValueOrDefault("DATA").Trim();
                        var nomeRespAtendimento = valoresLinha.GetValueOrDefault("NOME_RESP_ATEND").Trim();
                        var codRespAtendimento = valoresLinha.GetValueOrDefault("RESP_ATEND").Trim();
                        var numero = valoresLinha.GetValueOrDefault("NUMERO").Trim();
                        var documento = valoresLinha.GetValueOrDefault("DOCUMENTO").Trim();


                        cpf = cpf.ToCPF();
                        string? nome = null;

                        if (dataTablePessoas.Rows.Count > 0)
                        {
                            DataRow[] dataRowEncontrados = dataTablePessoas.AsEnumerable().Where(row => row.Field<string>("Documento(CPF,CNPJ,CGC)") == cpf).ToArray();
                            if (dataRowEncontrados.Length > 0)
                                nome = dataRowEncontrados[0]["NomeCompleto"].ToString();

                            //`Para analisar porque está vazio
                            if (!string.IsNullOrEmpty(nome))
                                continue;

                            DataRow[] dataRowRecebiveis = dataTableRecebiveis.AsEnumerable()
                            .Where(row =>
                                row.Field<string>("Documento Ref") == documento &&
                                row.Field<string>("CPF") == cpf)
                            .ToArray();

                            dataRow["Número do Controle"] = documento;
                            dataRow["Paciente CPF"] = cpf;
                            dataRow["Paciente Nome"] = nome;
                            dataRow["Dentista CPF"] = string.Empty;
                            dataRow["Dentista Nome"] = nomeRespAtendimento;
                            dataRow["Dente"] = numero;
                            dataRow["Procedimento Nome"] = nomeProduto;
                            //dataRow["Dentista Codigo"] = codRespAtendimento;
                            dataRow["Procedimento Valor"] = valor;
                            dataRow["Procedimento Observação"] = observacao;
                            dataRow["Data Início"] = dataInicio.ToData().ToShortDateString();
                            dataRow["Data Termino"] = dataAtendimento.ToDataNull();

                            dataTable.Rows.Add(dataRow);
                        }
                    }
                    catch (Exception error)
                    {
                        throw new Exception($"Erro na linha {linhaIndex + 1}: {error.Message}");
                    }

                    linhaIndex++;
                }

                return dataTable;
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao converter Excel Recebíveis: {error.Message}");
            }
        }
        public DataTable MergeDataTables(DataTable dt1, DataTable dt2)
        {
            //Merge dos Datatables Man001 e Man101 linkando ambos pelo cnpj

            var mergedDataTable = from row1 in dt1.AsEnumerable()
                                  join row2 in dt2.AsEnumerable() on row1.Field<string>("PacienteCPF") equals row2.Field<string>("PacienteCPF")
                                  select new
                                  {
                                      PacienteNomeCompleto001 = row1.Field<string>("PacienteNomeCompleto"),
                                      PacienteCPF001 = row1.Field<string>("PacienteCPF"),
                                      Observacao = row1.Field<string>("Observação"),
                                      DataModificado = row1.Field<string>("DataModificado"),
                                      Diagnostico = row1.Field<string>("Diagnostico"),
                                      DataInicial = row1.Field<string>("DataInicial"),
                                      DataFinal = row1.Field<string>("DataFinal"),
                                      PacienteCPF101 = row2.Field<string>("PacienteCPF"),
                                      PacienteNomeCompleto101 = row2.Field<string>("PacienteNomeCompleto"),
                                      DentistaCPF = row2.Field<string>("DentistaCPF"),
                                      DentistaNome = row2.Field<string>("DentistaNome"),
                                      DentistaCodigo = row2.Field<string>("DentistaCodigo"),
                                      Procedimento = row2.Field<string>("Procedimento"),
                                      DataAtendimento = row2.Field<string>("DataAtendimento"),
                                      DataInicio = row2.Field<string>("DataInicio"),
                                      DataRetorno = row2.Field<string>("DataRetorno")
                                  };

            DataTable resultDataTable = new DataTable();
            resultDataTable.Columns.Add("PacienteNomeCompleto001");
            resultDataTable.Columns.Add("PacienteCPF001");
            resultDataTable.Columns.Add("Observação");
            resultDataTable.Columns.Add("DataModificado");
            resultDataTable.Columns.Add("Diagnostico");
            resultDataTable.Columns.Add("DataInicial");
            resultDataTable.Columns.Add("DataFinal");
            resultDataTable.Columns.Add("PacienteCPF101");
            resultDataTable.Columns.Add("PacienteNomeCompleto101");
            resultDataTable.Columns.Add("DentistaCPF");
            resultDataTable.Columns.Add("DentistaNome");
            resultDataTable.Columns.Add("DentistaCodigo");
            resultDataTable.Columns.Add("Procedimento");
            resultDataTable.Columns.Add("DataAtendimento");
            resultDataTable.Columns.Add("DataInicio");
            resultDataTable.Columns.Add("DataRetorno");

            foreach (var item in mergedDataTable)
            {
                resultDataTable.Rows.Add(item.PacienteNomeCompleto001, item.PacienteCPF001, item.Observacao, item.Diagnostico,
                                         item.DataInicial, item.DataFinal, item.PacienteCPF101, item.PacienteNomeCompleto101,
                                         item.DentistaNome, item.DentistaCodigo, item.Procedimento, item.DataAtendimento,
                                         item.DataInicio, item.DataRetorno);
            }


            //Merge dos Datatables Man001 e Man101 sem linkar os cnpj

            //DataTable resultDataTable = new DataTable();          
            //foreach (DataColumn col in dt1.Columns)
            //{
            //    resultDataTable.Columns.Add(col.ColumnName);
            //}
            //foreach (DataColumn col in dt2.Columns)
            //{
            //    resultDataTable.Columns.Add(col.ColumnName);
            //}

            //for (int i = 0; i < dt1.Rows.Count; i++)
            //{
            //    DataRow item = resultDataTable.NewRow();
            //    foreach (DataColumn col in dt1.Columns)
            //    {
            //        item[col.ColumnName] = dt1.Rows[i][col.ColumnName];
            //    }
            //    foreach (DataColumn col in dt2.Columns)
            //    {
            //        item[col.ColumnName] = dt2.Rows[i][col.ColumnName];
            //    }
            //    resultDataTable.Rows.Add(item);
            //    resultDataTable.AcceptChanges();
            //}
            return resultDataTable;
        }
        public DataTable MergeDataTablesProcedimentosManutencao(DataTable dt1, DataTable dt2)
        {
            var mergedDataTable = from row1 in dt1.AsEnumerable()
                                      //join row2 in dt2.AsEnumerable() on row1.Field<string>("Paciente CPF") equals row2.Field<string>("Paciente CPF")
                                  select new
                                  {
                                      NumeroControle = row1.Field<string>("Número do Controle"),
                                      PacienteCPF = row1.Field<string>("Paciente CPF"),
                                      PacienteNome = row1.Field<string>("Paciente Nome"),
                                      DentistaCPF = row1.Field<string>("Dentista CPF"),
                                      DentistaNome = row1.Field<string>("Dentista Nome"),
                                      Dente = row1.Field<string>("Dente"),
                                      ProcedimentoNome = row1.Field<string>("Procedimento Nome"),
                                      ProcedimentoValor = row1.Field<string>("Procedimento Valor"),
                                      ProcObservacao = row1.Field<string>("Procedimento Observação"),
                                      DataInicio = row1.Field<string>("Data Início"),
                                      DataTermino = row1.Field<string>("Data Termino"),
                                      QtdOrto = string.Empty
                                      //ValorPagamento = row2.Field<string>("Valor Pagamento"),
                                      //DataPagamento = row2.Field<string>("Data Pagamento"),
                                  };   //).Distinct();            

            DataTable resultDataTable = new DataTable();
            resultDataTable.Columns.Add("Número do Controle");
            resultDataTable.Columns.Add("Paciente CPF");
            resultDataTable.Columns.Add("Paciente Nome");
            resultDataTable.Columns.Add("Dentista CPF");
            resultDataTable.Columns.Add("Dentista Nome");
            resultDataTable.Columns.Add("Dente");
            resultDataTable.Columns.Add("Procedimento Nome");
            resultDataTable.Columns.Add("Procedimento Valor");
            resultDataTable.Columns.Add("Procedimento Observação");
            resultDataTable.Columns.Add("Data Início");
            resultDataTable.Columns.Add("Data Termino");
            resultDataTable.Columns.Add("Quantidade Orto");

            //resultDataTable.Columns.Add("Valor Pagamento");
            //resultDataTable.Columns.Add("Data Pagamento");


            foreach (var item in mergedDataTable)
            {
                resultDataTable.Rows.Add(item.NumeroControle, item.PacienteCPF, item.PacienteNome, item.DentistaCPF,
                                         item.DentistaNome, item.Dente, item.ProcedimentoNome, item.ProcedimentoValor,
                                         item.ProcObservacao, item.DataInicio, item.DataTermino, item.QtdOrto);
            }

            var mergedDataTable2 = from row2 in dt2.AsEnumerable()
                                   select new
                                   {
                                       NumeroControle = string.Empty,
                                       PacienteCPF = row2.Field<string>("Paciente CPF"),
                                       PacienteNome = row2.Field<string>("Paciente Nome"),
                                       DentistaCPF = string.Empty,
                                       DentistaNome = row2.Field<string>("Dentista Nome"),
                                       Dente = row2.Field<string>("Dente"),
                                       ProcedimentoNome = row2.Field<string>("Procedimento Nome"),
                                       ProcedimentoValor = row2.Field<string>("Procedimento Valor"),
                                       ProcObservacao = row2.Field<string>("Procedimento Observação"),
                                       DataInicio = string.Empty,
                                       DataTermino = string.Empty,
                                       QtdOrto = row2.Field<string>("Quantidade Orto"),
                                       TipoPagamento = row2.Field<string>("Tipo do Pagamento"),
                                       Vencimento = row2.Field<string>("Vencimento"),
                                       ValorDevido = row2.Field<string>("Valor Devido")
                                        //ValorPagamento = row2.Field<string>("Valor Pagamento"),
                                        //DataPagamento = row2.Field<string>("Data Pagamento"),
        };

            foreach (var item in mergedDataTable2)
            {
                resultDataTable.Rows.Add(item.NumeroControle, item.PacienteCPF, item.PacienteNome, item.DentistaCPF,
                                         item.DentistaNome, item.Dente, item.ProcedimentoNome, item.ProcedimentoValor,
                                         item.ProcObservacao, item.DataInicio, item.DataTermino, item.QtdOrto,
                                         item.TipoPagamento, item.Vencimento,item.ValorDevido);
            }

            return resultDataTable;
        }
        public void RetornaProcedimentosPorTipo(DataTable dt, string estabelecimentoID)
        {
            ExcelHelper excelHelper = new();
            DataTable dataTableProcedimentos = new();
            dataTableProcedimentos.Columns.Add("Nome Tabela");
            dataTableProcedimentos.Columns.Add("Especialidade");
            dataTableProcedimentos.Columns.Add("tivo (Sim/Não)");
            dataTableProcedimentos.Columns.Add("Nome do Procedimento");
            dataTableProcedimentos.Columns.Add("Abreviação");
            dataTableProcedimentos.Columns.Add("Preço");
            dataTableProcedimentos.Columns.Add("TUSS");
            dataTableProcedimentos.Columns.Add("Especialidade Código");

            #region DataRows pormtipo de plano

            DataRow[] dataRowParticular = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Particular").ToArray();
            DataRow[] dataRowOdc = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Odc").ToArray();
            DataRow[] dataRowAmil = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Amil").ToArray();
            DataRow[] dataRowOdontomaxi = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Odontomaxi").ToArray();
            DataRow[] dataRowUnimed = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Unimed").ToArray();
            DataRow[] dataRowPrimavida = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Primavida").ToArray();
            DataRow[] dataRowPortoSeguro = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Porto Seguro").ToArray();
            DataRow[] dataRowAesp = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Aesp").ToArray();
            DataRow[] dataRowRodriguesLeira = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Rodrigues Leira").ToArray();
            DataRow[] dataRowInpao = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Inpao").ToArray();
            DataRow[] dataRowDentalIntegral = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Dental Integral").ToArray();
            DataRow[] dataRowProasa = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Proasa").ToArray();
            DataRow[] dataRowIdealOdonto = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Ideal Odonto").ToArray();
            DataRow[] dataRowOdontoart = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Odontoart").ToArray();
            DataRow[] dataRowBrazilDental = dt.AsEnumerable().Where(row => row.Field<string>("Nome Tabela") == "Brazil Dental").ToArray();

            #endregion

            #region Particular

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowParticular)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosParticular = Tools.GerarNomeArquivo($"CadastroProcedimentos_Particular_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosParticular + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Odc

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowOdc)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosOdc = Tools.GerarNomeArquivo($"CadastroProcedimentos_Odc_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosOdc + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Amil

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowAmil)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosAmil = Tools.GerarNomeArquivo($"CadastroProcedimentos_Amil_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosAmil + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Odontomaxi

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowOdontomaxi)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosOdontomaxi = Tools.GerarNomeArquivo($"CadastroProcedimentos_Odontomaxi_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosOdontomaxi + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Unimed

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowUnimed)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosUnimed = Tools.GerarNomeArquivo($"CadastroProcedimentos_Unimed_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosUnimed + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Primavida

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowPrimavida)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosPrimavida = Tools.GerarNomeArquivo($"CadastroProcedimentos_Primavida_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosPrimavida + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Porto Seguro

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowPortoSeguro)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosPortoSeguro = Tools.GerarNomeArquivo($"CadastroProcedimentos_PortoSeguro_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosPortoSeguro + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Aesp

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowAesp)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosAesp = Tools.GerarNomeArquivo($"CadastroProcedimentos_Aesp_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosAesp + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Rodrigues Leira

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowRodriguesLeira)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosRodriguesLeira = Tools.GerarNomeArquivo($"CadastroProcedimentos_RodriguesLeira_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosRodriguesLeira + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Inpao

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowInpao)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosInpao = Tools.GerarNomeArquivo($"CadastroProcedimentos_Inpao_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosInpao + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Dental Integral

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowDentalIntegral)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosDentalIntegral = Tools.GerarNomeArquivo($"CadastroProcedimentos_DentalIntegral_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosDentalIntegral + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Proasa

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowProasa)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosProasa = Tools.GerarNomeArquivo($"CadastroProcedimentos_Proasa_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosProasa + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Ideal Odonto

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowIdealOdonto)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosIdealOdonto = Tools.GerarNomeArquivo($"CadastroProcedimentos_IdealOdonto_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosIdealOdonto + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Odontoart

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowOdontoart)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosOdontoart = Tools.GerarNomeArquivo($"CadastroProcedimentosOdontoart_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosOdontoart + ".xlsx", dataTableProcedimentos);
            }

            #endregion

            #region Brazil Dental

            dataTableProcedimentos?.Clear();

            foreach (var x in dataRowBrazilDental)
                dataTableProcedimentos?.Rows.Add(x.ItemArray);

            if (dataTableProcedimentos != null)
            {
                var salvarProcedimentosBrazilDental = Tools.GerarNomeArquivo($"CadastroProcedimentos_BrazilDental_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarProcedimentosBrazilDental + ".xlsx", dataTableProcedimentos);
            }

            #endregion
        }
    }
}
