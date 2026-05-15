using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class InventarioCargaSoltaArmazemViewModel
    {
        public int PATIO { get; set; }
        public int COD_USUARIO { get; set; }
    }

    public class InventarioCsComboDto
    {
        public int AUTONUM { get; set; }
        public string DISPLAY { get; set; }
    }

    public class InventarioCsItemLinhaDto
    {
        /// <summary>Valor usado no combo (equivalente ao BoundText do VB: prefixo C/Y/R/T + identificador).</summary>
        public string ITEM_CHAVE { get; set; }
        public string DISPLAY { get; set; }
    }

    public class InventarioCsHistoricoLinhaDto
    {
        public string ARMAZEM { get; set; }
        public string LOCAL { get; set; }
        public string DATA_MOV { get; set; }
        public string USUARIO { get; set; }
    }

    public class InventarioCsMarcanteEstadoResponse
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public bool PATIO_DIVERGENTE { get; set; }
        public string MODO_TAG { get; set; }
        public string TITULO_STATUS { get; set; }
        public bool FLAG_EXPORTACAO { get; set; }
        public bool FLAG_RECONHECIMENTO { get; set; }

        public string LOTE { get; set; }
        public string LOTE_TAG { get; set; }
        public string MERCADORIA { get; set; }
        public string MARCA { get; set; }
        public string ENTRADA { get; set; }
        public string CONTEINER { get; set; }
        public string CLIENTE { get; set; }
        public string EMBALAGEM { get; set; }
        public string DOC { get; set; }
        public string CANAL { get; set; }
        public string MOV { get; set; }
        public string IMO { get; set; }
        public string NVOCC { get; set; }
        public string BL { get; set; }
        public string LOCAL_ATUAL { get; set; }
        public int QTDE { get; set; }
        public int QTDE_M { get; set; }
        public int QTDE_POS { get; set; }
        public long AUTONUM_CS_YARD { get; set; }
        public int? MOTIVO_PADRAO { get; set; }

        public List<InventarioCsItemLinhaDto> ITENS { get; set; } = new List<InventarioCsItemLinhaDto>();
    }

    public class InventarioCsItemDetalheResponse
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public int QTDE { get; set; }
        public string EMBALAGEM { get; set; }
        public string LOCAL { get; set; }
        public int? MOTIVO_PADRAO { get; set; }
    }

    public class InventarioCsResolverPrateleiraRequest
    {
        public string TEXTO { get; set; }
    }

    public class InventarioCsResolverPrateleiraResponse
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public int ARMAZEM_ID { get; set; }
        public string YARD { get; set; }
        public string YARD_TAG { get; set; }
        public bool MODO_CONTENER { get; set; }
    }

    public class InventarioCsSalvarRequest
    {
        public int MARCANTE { get; set; }
        public string LOTE { get; set; }
        public string LOTE_TAG { get; set; }
        public string ITEM_CHAVE { get; set; }
        public string MODO_TAG { get; set; }
        public bool FLAG_EXPORTACAO { get; set; }
        public bool FLAG_MODO_CONTENER { get; set; }
        public int ARMAZEM_ID { get; set; }
        public string YARD { get; set; }
        public string YARD_TAG { get; set; }
        public int QUANTIDADE_POS { get; set; }
        public int QTDE_DISPONIVEL { get; set; }
        public int QTDE_MARCANTE { get; set; }
        public int MOTIVO_ID { get; set; }
        public string YARD_ATUAL_TEXTO { get; set; }
        public int FL_FRENTE { get; set; }
        public int FL_FUNDO { get; set; }
        public int FL_LE { get; set; }
        public int FL_LD { get; set; }
        public int PERCENT_CT { get; set; }
        /// <summary>IMO exibido na tela (para futura segregacao pos-salvar, espelho do legado).</summary>
        public string IMO { get; set; }
    }

    public class InventarioCsSalvarResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public string MARCANTE_SUGERIDO_CONFLITO { get; set; }
    }
}
