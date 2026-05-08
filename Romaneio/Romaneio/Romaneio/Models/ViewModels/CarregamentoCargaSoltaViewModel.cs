using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class CarregamentoCargaSoltaViewModel
    {
        public int PATIO { get; set; }
        public bool BIPAGEM_OBRIGATORIA { get; set; }
        public bool BIPAGEM_BLOQUEAR_APOS_VALIDAR { get; set; }
        public List<VeiculoCarregamentoDto> VEICULOS { get; set; } = new List<VeiculoCarregamentoDto>();
    }

    public class VeiculoCarregamentoDto
    {
        public string PLACA_C { get; set; }
        public string PLACA_CARRETA { get; set; }
        public string MODELO { get; set; }
        public string DISPLAY { get; set; }
    }

    public class OrdemCarregamentoDto
    {
        public int NUM_OC { get; set; }
        public int LOTE { get; set; }
        public int ITEM { get; set; }
        public int QUANTIDADE { get; set; }
        public int QTDE_CARREGADA { get; set; }
        public string EMBALAGEM { get; set; }
        public int ORDEM_CARREG { get; set; }
        public int AUTONUMCS { get; set; }
    }

    public class MarcanteCarregadoDto
    {
        public int MARCANTE { get; set; }
        public int QTDE { get; set; }
        public string MARCANTE_FORMATADO => MARCANTE.ToString("000000000");
    }

    public class LocalArmazemDto
    {
        public int MARCANTE { get; set; }
        public int QTDE { get; set; }
        public string DESCR_ARMAZEM { get; set; }
        public string POSICAO { get; set; }
        public string MARCANTE_FORMATADO => MARCANTE.ToString("000000000");
        public string LOCAL_COMPLETO => string.Format("{0} {1}", DESCR_ARMAZEM ?? string.Empty, POSICAO ?? string.Empty).Trim();
    }

    public class MarcanteDadosDto
    {
        public int MARCANTE { get; set; }
        public int AUTONUMCS { get; set; }
        public int LOTE { get; set; }
        public int ITEM { get; set; }
        public int VOLUMES { get; set; }
        public string EMBALAGEM { get; set; }
        public string MERCADORIA { get; set; }
        public string MARCA { get; set; }
        public int AUTONUM_ARMAZEM { get; set; }
        public string DESCR_ARMAZEM { get; set; }
        public int? AUTONUMCNTR { get; set; }
        public string ID_CONTEINER { get; set; }
        public string POSICAO_YARD { get; set; }
        public int AUTONUM_CS_YARD { get; set; }
    }

    public class CarregarItemRequest
    {
        public int AUTONUMCS_OC { get; set; }
        public int ORDEM { get; set; }
        public string PLACA_C { get; set; }
        public int MARCANTE { get; set; }
        public int QTDE { get; set; }
        public int AUTONUM_CS_YARD { get; set; }
        public int AUTONUM_ARMAZEM { get; set; }
    }

    public class CarregarItemResult
    {
        public bool SUCESSO { get; set; }
        public bool PARCIAL { get; set; }
        public string MENSAGEM { get; set; }
    }

    public class CancelarItemRequest
    {
        public int MARCANTE { get; set; }
        public int AUTONUMCS_OC { get; set; }
        public int AUTONUM_CS_YARD { get; set; }
        public int AUTONUM_ARMAZEM { get; set; }
        public int QTDE { get; set; }
    }

    public class CancelarItemResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
    }

    public class DocumentoOcResult
    {
        public bool SUCESSO { get; set; }
        public bool LTL { get; set; }
        public int? AUTONUM_AG_CS { get; set; }
        public int? LOTE_BL { get; set; }
        public string MENSAGEM { get; set; }
    }
}
