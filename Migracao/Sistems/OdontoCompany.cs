using Migracao.Models;
using Migracao.Utils;
using NPOI.SS.Formula.Functions;
using System.Data;
using System.Text;
using static System.Windows.Forms.LinkLabel;

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
        string[] MAN001_DesenvClinico = ["CNPJ_CPF","QTDE_MANUT","OBS","ATIVO","INICIO_MANUT","FINAL_MANUT","MODIFICADO","USUARIO","DATA_ATIVO","USU_ATIVO","USU_INSTALACAO","USU_RETIRADA","MOTIVO_ATIVO","DIAGNOSTICO","PROGNOSTICO","OBS_CLASSE","CLASSE","DATA_ALTERACAO","HORA_ALTERACAO","USUARIO_ALTERACAO","DATA_MODIFICADO","DT_AXON","AXON_ID"
];

        List<string> cabecalhos_Pacientes = ["Código", "Ativo(S/N)", "NomeCompleto", "NomeSocial", "Apelido", "Documento(CPF,CNPJ,CGC)", "DataCadastro(01/12/2024)", "Observações", "Email", "RG", "Sexo(M/F)", "NascimentoData", "NascimentoLocal", "EstadoCivil(S/C/V)", "Profissao", "CargoNaClinica", "Dentista(S/N)", "ConselhoCodigo", "Paciente(S/N)", "Funcionario(S/N)", "Fornecedor(S/N)", "TelefonePrincipal", "Celular", "TelefoneAlternativo", "Logradouro", "LogradouroNum", "Complemento", "Bairro", "Cidade", "Estado(SP)", "CEP(00000-000)"];
        List<string> cabecalhos_Recebiveis = ["CPF", "Nome", "DocumentoRef", "RecebívelExigível(R/E)", "ValorOriginal", "ValorPago", "Prazo", "Vencimento(01/12/2010)", "DataBaixa", "Emissão(01/12/2010)", "ObservaçãoRecebível", "ObservaçãoRecebido"];
        List<string> cabecalhos_Agendamentos = ["ID", "CPF", "Nome Completo", "Telefone", "Data Início (01/12/2024 00:00)", "Data Término (01/12/2024 00:00)", "Data Inclusão (01/12/2024)", "NomeCompletoDentista", "Observacao"];
        List<string> cabecalhos_Procedimentos = ["Nome Tabela", "Ativo(S/N)", "Procedimento(Nome)", "Abreviação", "Especialidade", "Especialidade Código", "Especialidade ID", "Preço", "TUSS", "Diagnóstico(S/N)", "Prevenção(S/N)", "Odontopediatria(S/N)", "Dentística(S/N)", "Endodontia(S/N)", "Periodontia(S/N)", "Prótese(S/N)", "Cirurgia(S/N)", "Ortodontia(S/N)", "Radiologia(S/N)", "Estética(S/N)", "Implantodontia(S/N)", "Odontogeriatria(S/N)", "DTM(S/N)", "Orofacial(S/N)",];
        List<string> cabecalhos_CodProcedimentos = ["ID", "Nome", "Usuário"];
        List<string> cabecalhos_DesenvClinico = ["Nome", "CNPJ_CPF", "Observação", "DataModificado", "Diagnostico", "ObservaçãoClasse"];

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
            DataTable dataTableDesenvClinico = new();
            //registroRecebivel = new HashSet<string>();

            foreach (string coluna in cabecalhos_Pacientes)
                dataTablePessoas.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_Recebiveis)
                dataTableRecebiveis.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_Agendamentos)
                dataTableAgendamentos.Columns.Add(coluna, typeof(string));

            foreach (string coluna in cabecalhos_DesenvClinico)
                dataTableDesenvClinico.Columns.Add(coluna, typeof(string));

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

                dataTableRecebiveis = ConvertExcelRecebiveisHistoricoVendas(dataTableRecebiveis, cabecalhosCSV, lstVendaCSV, dataTablePessoas);

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
            }

            var excel_MAN001 = listView.Items.Cast<ListViewItem>()
                .FirstOrDefault(item => item.SubItems.Cast<ListViewItem.ListViewSubItem>().Any(s => s.Text.Contains("MAN001")));

            if (excel_MAN001 != null )
            {
                var resultado = LerArquivosExcelCsv(excel_MAN001.Text, Encoding.UTF8);
                var linhasCSV = resultado.Item1;
                var cabecalhosCSV = resultado.Item2;

                if (MAN001_DesenvClinico.All(cabecalhosCSV.Contains))
                    dataTableDesenvClinico = ConvertExcelDesenvolvimentoClinico(dataTableDesenvClinico, cabecalhosCSV, linhasCSV, dataTablePessoas);
            }

            if (dataTableProcedimentos.Rows.Count > 0)
            {
                var salvarArquivoAgenda = Tools.GerarNomeArquivo($"CadastroProcedimentos_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoAgenda + ".xlsx", dataTableProcedimentos);
            }

            if (excel_BXD111 != null || excel_CRD111 != null || excel_CXD555 != null)
            {
                var salvarArquivoRecebiveis = Tools.GerarNomeArquivo($"CadastroRecebiveis_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarArquivoRecebiveis + ".xlsx", dataTableRecebiveis);
            }

            if(excel_MAN001 != null)
            {
                var salvarDesenvolvimentoCLinico = Tools.GerarNomeArquivo($"CadastroDesenvolvimentoCLinico_{estabelecimentoID}_OdontoCompany");
                excelHelper.CriarExcelArquivo(salvarDesenvolvimentoCLinico + ".xlsx", dataTableDesenvClinico);
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
                            dataRow["ObservaçãoRecebível"] = observacao;
                            dataRow["DocumentoRef"] = documento;
                            dataRow["ValorOriginal"] = valor.ArredondarValorV2();
                            //dataRow["Prazo"] = prazo;
                            dataRow["Vencimento(01/12/2010)"] = vencimentoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["Emissão(01/12/2010)"] = emissaoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["RecebívelExigível(R/E)"] = "R";

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
        public DataTable ConvertExcelRecebiveisHistoricoVendas
            (DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePacientes)
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
                            dataRow["ObservaçãoRecebível"] = observacao;
                            dataRow["DocumentoRef"] = documento;
                            dataRow["ValorOriginal"] = valor.ArredondarValorV2();
                            //dataRow["Prazo"] = prazo;
                            dataRow["Vencimento(01/12/2010)"] = vencimentoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["Emissão(01/12/2010)"] = emissaoData.ToData().ToString("dd/MM/yyyy");
                            dataRow["RecebívelExigível(R/E)"] = "R";

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
                        DataRow dataRow = dataTable.NewRow();
                        var valoresLinha = new Dictionary<string, string>();

                        for (int i = 0; i < cabecalhos.Count; i++)
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
                                row.Field<string>("DocumentoRef") == documento &&
                                row.Field<string>("CPF") == cpf &&
                                row.Field<string>("ValorOriginal") == valorOriginal.ArredondarValorV2().ToString())
                            .ToArray();

                            if (dataRowEncontrados.Length > 0)
                            {
                                foreach (var dataRowEncontrado in dataRowEncontrados)
                                {
                                    dataRowEncontrado["ValorPago"] = valor.ArredondarValorV2();
                                    dataRowEncontrado["DataBaixa"] = baixaData.ToData().ToString("dd/MM/yyyy");
                                    dataRowEncontrado["ObservaçãoRecebido"] = observacao;
                                }
                            }

                            else
                            {
                                dataRow["CPF"] = cpf;
                                dataRow["Nome"] = nome;
                                dataRow["DocumentoRef"] = documento;
                                dataRow["ValorPago"] = valor.ArredondarValorV2();
                                dataRow["DataBaixa"] = baixaData;
                                dataRow["Vencimento(01/12/2010)"] = vencimentoData.ToData();
                                dataRow["ObservaçãoRecebido"] = observacao;
                                dataRow["RecebívelExigível(R/E)"] = "R";

                                dataTable.Rows.Add(dataRow);
                            }
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

                        dataRow["Nome Tabela"] = nomeTabela.GetPrimeirosCaracteres(40).PrimeiraLetraMaiuscula();
                        dataRow["Especialidade"] = especialidade;
                        dataRow["Ativo(S/N)"] = ativo == "N" ? "N" : "S";
                        dataRow["Procedimento(Nome)"] = nome.GetPrimeirosCaracteres(100).PrimeiraLetraMaiuscula();
                        dataRow["Abreviação"] = abreviacao;
                        dataRow["Preço"] = valor.ArredondarValorV2();
                        dataRow["TUSS"] = tuss.ToNumV2();
                        dataRow["Especialidade Código"] = especialidadeCod;
                        dataRow["Especialidade ID"] = especialidadeCodID;

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

        public DataTable ConvertExcelDesenvolvimentoClinico(DataTable dataTable, List<string> cabecalhos, List<string[]> linhas, DataTable dataTablePessoas = null)
        {
            try
            {
                int linhaIndex = 0;
                string dataModificadoTeste = null;

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
    }
}
