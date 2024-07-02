﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class FinanceiroRecebidosDTO
    {
        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("Nome")]
        public string? Nome { get; set; }

        [DisplayName("Número do Controle")]
        public string? Numero_Controle { get; set; }

        [DisplayName("Recebível Exigível(R/E)")]
        public string? Recebivel_Exigivel { get; set; }

        [DisplayName("Valor Devido")]
        public string? Valor_Devido { get; set; }

        [DisplayName("Valor Pago")]
        public string? Valor_Pago { get; set; }

        [DisplayName("Prazo")]
        public string? Prazo { get; set; }

        [DisplayName("Data Vencimento")]
        public string? Data_Vencimento { get; set; }

        [DisplayName("Data do Pagamento")]
        public string? Data_Pagamento { get; set; }

        [DisplayName("Emissão")]
        public string? Emissao { get; set; }

        [DisplayName("Observação Recebido")]
        public string? Observacao_Recebido { get; set; }

        [DisplayName("Tipo Pagamento")]
        public string? Tipo_Pagamento { get; set; }

        [DisplayName("Valor Original")]
        public string? Valor_Original { get; set; }

        [DisplayName("Vencimento Recebível")]
        public string? Vencimento_Recebivel { get; set; }

        [DisplayName("Duplicata")]
        public string? Duplicata { get; set; }

        [DisplayName("Parcela")]
        public string? Parcela { get; set; }

        [DisplayName("Tipo Espécie Pagamento")]
        public string? Tipo_Especie_Pagamento { get; set; }

        [DisplayName("Espécie Pagamento")]
        public string? Especie_Pagamento { get; set; }
    }
}
