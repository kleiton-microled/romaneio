using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class LocalizacaoCarregamentosViewModel
    {
        public int PATIO { get; set; }
        public string DESCR_PATIO { get; set; }
    }

    public class LocalizacaoCarregamentosFiltroRequest
    {
        public string PLACA { get; set; }
    }

    public class LocalizacaoCarregamentosFiltroResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public List<LocalizacaoCarregamentoItemDto> ITENS { get; set; } = new List<LocalizacaoCarregamentoItemDto>();
        public int TOTAL_ENTRADAS { get; set; }
        public int TOTAL_CNTR { get; set; }
        public int TOTAL_CS { get; set; }
    }

    public class LocalizacaoCarregamentoItemDto
    {
        public string PLACA { get; set; }
        public string CARRETA { get; set; }
        public string DT_LIB_ENT_CAM { get; set; }
        public string DT_GATE_IN { get; set; }
        public string CARGA { get; set; }
        public string QUADRA { get; set; }
        public string RUA { get; set; }
        public string NOMETRANSPORTADORA { get; set; }
        public string USUARIO { get; set; }
        public int FLAG_CNTR { get; set; }
        public int AUTONUM_CNTR { get; set; }
    }

    public class MoverSaidaRequest
    {
        public int AUTONUM_CNTR { get; set; }
        public string QUADRA_ATUAL { get; set; }
        public int FLAG_CNTR { get; set; }
    }

    public class MoverSaidaResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
    }
}
