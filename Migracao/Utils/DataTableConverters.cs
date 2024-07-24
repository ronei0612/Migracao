using Migracao.DTO;
using Migracao.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Utils
{
    public class DataTableConverters
    {
        public static DataTable ConversorEntidadeParaDataTable(List<PacientesDentistasDTO> pacientesDTO)
        {
            var dataTable = new DataTable();

            // Adiciona as colunas ao DataTable baseado nos nomes das propriedades da classe Person
            foreach (var prop in typeof(PacientesDentistasDTO).GetProperties())
            {
                dataTable.Columns.Add(prop.Name, prop.PropertyType);
            }

            // Usando Parallel.ForEach para processar a lista de pessoas e preencher o DataTable
            Parallel.ForEach(pacientesDTO, new ParallelOptions { MaxDegreeOfParallelism = 8 }, paciente =>
            {
                try
                {
                    DataRow row;

                    lock (new object())
                    { row = dataTable.NewRow(); }


                    // Preenche as células da linha com os valores das propriedades da pessoa
                    foreach (var prop in typeof(PacientesDentistasDTO).GetProperties())
                    {
                        row[prop.Name] = prop.GetValue(paciente);
                    }

                    // Adiciona a linha ao DataTable de forma thread-safe
                    lock (dataTable)
                    {
                        dataTable.Rows.Add(row);
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }

            });

            return dataTable;
        }
    }
}
