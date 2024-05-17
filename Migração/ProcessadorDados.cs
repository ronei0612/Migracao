using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Migração
{
	internal class ProcessadorDados
	{
		private string mascaraCPF;
		private string celulaValor;

		public ProcessadorDados(string celulaValor)
		{
			mascaraCPF = "000.000.000-00";
			this.celulaValor = celulaValor;
		}

		public string ProcessarMascara()
		{
			mascaraCPF = mascaraCPF.Split('.')[0].Replace(".", @"\.").Replace("-", @"\-");
			var mascaraCPFLenth = Regex.Replace(mascaraCPF, "[^0-9]", "").Length.ToString();
			return mascaraCPFLenth;
		}

		public string ObterNomeCompleto()
		{
			var nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
			return nomeCompleto;
		}
	}
}
