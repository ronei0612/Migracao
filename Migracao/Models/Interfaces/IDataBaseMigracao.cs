﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.Interfaces
{
    public interface IDataBaseMigracao
    {
        void DataBaseImportacaoProcedimentos();

        void DataBaseImportacaoDevClinico();

        void DataBaseImportacaoProntuarios();

        void DataBaseImportacaoManutencoes();

        void DataBaseImportacaoPagosExigiveis();

        void DataBaseImportacaoFinanceiroRecebiveis();

        void DataBaseImportacaoPacientesDentistas();

        void DataBaseImportacaoAgendamentos();

        void DataBaseImportacaoDentistas();

        void DataBaseImportacaoRecebiveisHistVenda();

        void DataBaseImportacaoProcedimentosPrecos();
    }
}
