﻿using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Text.RegularExpressions;

namespace Migracao.Utils
{
    public class ConverterHelper
    {
        public void JsonExcel(string json, string caminhoArquivoExcel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            json = File.ReadAllText(json);
            // Converte o JSON em uma lista de dicionários
            var dados = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);

            // Cria um DataTable a partir dos dados do JSON
            var dataTable = new DataTable();
            foreach (var key in dados[0].Keys)
            {
                dataTable.Columns.Add(key);
            }
            foreach (var item in dados)
            {
                try
                {
                    var row = dataTable.NewRow();
                    foreach (var key in item.Keys)
                    {
                        row[key] = item[key];
                    }
                    dataTable.Rows.Add(row);
                }
                catch { }
            }

            // Cria um arquivo Excel a partir do DataTable
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                package.SaveAs(new FileInfo(caminhoArquivoExcel));
            }
        }

        public static List<string> SomaValoresColunaPorReferencia(DataTable dt, string nomeColunaReferencia, string nomeColuna)
        {
            var query = (from row in dt.AsEnumerable()
                         group row by row.Field<string>(nomeColunaReferencia) into g
                         select new
                         {
                             Soma = g.Sum(row => row.Field<double>(nomeColuna)).ToString()
                         }).ToList();

            return null;
        }
    }
}
