using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class ManutencaoMarcantesArmazemViewModel
    {
    }

    public class ManutencaoMarcantesArmazemLinhaDto
    {
        public int Indice { get; set; }

        public string DescrArmazem { get; set; }

        public string Posicao { get; set; }

        public int? Item { get; set; }

        public int Qtde { get; set; }

        public int Marcante { get; set; }

        public string MarcanteFormatado => Marcante.ToString("000000000");

        public int AutonumCs { get; set; }
    }

    public class ManutencaoMarcantesArmazemFiltrarResult
    {
        public string MensagemErro { get; set; }

        public string Embalagem { get; set; }

        public int? QuantidadeReferencia { get; set; }

        public int? AutonumCs { get; set; }

        public List<ManutencaoMarcantesArmazemLinhaDto> Linhas { get; set; } = new List<ManutencaoMarcantesArmazemLinhaDto>();

        public int TotalVolumesListados { get; set; }

        public bool TotalConfereReferencia { get; set; }
    }

    public class ManutencaoMarcantesArmazemSalvarRequest
    {
        public int AutonumCs { get; set; }

        public int MarcanteAutonum { get; set; }

        /// <summary>Null ou ausente equivale a campo vazio no VB6.</summary>
        public int? NovaQuantidade { get; set; }

        /// <summary>
        /// Quantidade de referencia da divergencia (mesmo valor exibido apos o filtro), para recalcular o indicador visual.
        /// </summary>
        public int? QuantidadeReferencia { get; set; }

        /// <summary>Cabecalho exibido (eco do filtro) caso a view nao permita recarregar por AUTONUM_CS.</summary>
        public string Embalagem { get; set; }
    }
}
