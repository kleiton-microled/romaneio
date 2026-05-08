using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class ConsultaLiberacaoCarregamentoViewModel
    {
        public List<TipoVeiculoDto> TIPOS_VEICULO { get; set; } = new List<TipoVeiculoDto>();
    }

    public class ConsultaLiberacaoFiltroRequest
    {
        public string PROTOCOLO { get; set; }
        public string PLACA { get; set; }
        public string CNTR { get; set; }
        public string BL { get; set; }
    }

    public class ConsultaLiberacaoResultadoDto
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public bool PODE_REGISTRAR { get; set; }
        public bool BLOQUEIO_BL { get; set; }
        public bool BLOQUEIO_CNTR { get; set; }
        public bool GLME_PENDENTE { get; set; }
        public int? ORDEM_CARREGAMENTO { get; set; }
        public int? AUTONUM_LOTE { get; set; }
        public int? AUTONUM_CNTR { get; set; }
        public int? AUTONUM_CS { get; set; }
        public int? TIPO_CAMINHAO { get; set; }
        public string PROTOCOLO { get; set; }
        public string PERIODO { get; set; }
        public string LOTE { get; set; }
        public string NUM_DOCUMENTO { get; set; }
        public string TRANSPORTADORA { get; set; }
        public string MOTORISTA { get; set; }
        public string PLACAS { get; set; }
        public string CPF_MOTORISTA { get; set; }
        public string STATUS_SISCARGA { get; set; }
        public string STATUS_SEFAZ { get; set; }
        public string LOCALIZACAO { get; set; }
        public List<CargaListaItemDto> CARGAS { get; set; } = new List<CargaListaItemDto>();
        public List<string> PROTOCOLOS_SIMULTANEOS { get; set; } = new List<string>();
    }

    public class CargaListaItemDto
    {
        public string CARGA { get; set; }
    }

    public class TipoVeiculoDto
    {
        public int AUTONUM { get; set; }
        public string DISPLAY { get; set; }
    }

    public class RegistrarLiberacaoRequest
    {
        public int ORDEM_CARREGAMENTO { get; set; }
        public int AUTONUM_LOTE { get; set; }
        public int? AUTONUM_CNTR { get; set; }
        public int? AUTONUM_CS { get; set; }
        public int TIPO_CAMINHAO { get; set; }
        public string PLACA_CAVALO { get; set; }
        public string PLACA_CARRETA { get; set; }
        public string CPF_MOTORISTA { get; set; }
    }

    public class RegistrarLiberacaoResultDto
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
    }
}
