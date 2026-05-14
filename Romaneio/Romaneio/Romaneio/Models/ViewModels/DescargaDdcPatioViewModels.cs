using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class DescargaDdcPatioViewModel
    {
        public int PATIO { get; set; }
        public string DESCR_PATIO { get; set; }
    }

    public class PlacaDdcListItemDto
    {
        public long AUTONUM_RS { get; set; }
        public string PLACA { get; set; }
    }

    public class LoteAtivoDto
    {
        public long AUTONUM_LOTE { get; set; }
    }

    public class GateLotesResultDto
    {
        public long GATE { get; set; }
        public List<LoteAtivoDto> LOTES { get; set; }
    }

    public class CarregarCabecalhoResultDto
    {
        public long CNTR { get; set; }
        public long GATE { get; set; }
        public long ID_OC { get; set; }
        public long LOTE { get; set; }
    }

    public class DescargaArmazemItemDto
    {
        public decimal QUANTIDADE { get; set; }
        public string DESCR_EMB { get; set; }
        public int FINALIZADO { get; set; }
        public long AUTONUM { get; set; }
        public string EMBALAGEM { get; set; }
    }

    public class EmbalagemDto
    {
        public string CODE { get; set; }
        public string DESCR { get; set; }
    }

    public class SalvarDescargaDdcRequest
    {
        public long AutonumRegistroSaida { get; set; }
        public long Cntr { get; set; }
        public long Gate { get; set; }
        public long Lote { get; set; }
        public long IdOc { get; set; }
        public decimal Quantidade { get; set; }
        public string Embalagem { get; set; }
        /// <summary>0 = NAO, 1 = SIM (conforme OptionButton do legado).</summary>
        public int Finalizado { get; set; }
        public long AutonumItem { get; set; }
    }

    public class DescargaDdcPatioOperacaoResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public string AVISO { get; set; }
    }
}
