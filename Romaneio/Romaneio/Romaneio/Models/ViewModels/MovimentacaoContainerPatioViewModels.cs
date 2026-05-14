using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class MovimentacaoContainerPatioViewModel
    {
        public int PATIO { get; set; }
        public string DESCR_PATIO { get; set; }
        public List<MotivoMovimentacaoDto> MOTIVOS { get; set; } = new List<MotivoMovimentacaoDto>();
        public List<CameraPatioDto> CAMERAS { get; set; } = new List<CameraPatioDto>();
        public MovimentacaoContainerPatioPermissoesDto PERMISSOES_AVARIAS { get; set; } = new MovimentacaoContainerPatioPermissoesDto();
    }

    public class MovimentacaoContainerPatioPermissoesDto
    {
        public bool ABA_AVARIAS_HABILITADA { get; set; }
        public bool PODE_EXCLUIR_AVARIA { get; set; }
        public bool PODE_INCLUIR_AVARIA { get; set; }
        public bool PODE_FINALIZAR_AVARIA { get; set; }
    }

    public class MotivoMovimentacaoDto
    {
        public int AUTONUM { get; set; }
        public string DESCRICAO { get; set; }
        public bool FLAG_CAMERA { get; set; }
    }

    public class CameraPatioDto
    {
        public int AUTONUM { get; set; }
        public string DESCR { get; set; }
    }

    public class CntrPatioBuscaFinalResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public string ID_CONTEINER { get; set; }
    }

    public class CntrPatioDetalheDto
    {
        public long AUTONUM { get; set; }
        public string ID_CONTEINER { get; set; }
        public string SISTEMA { get; set; }
        public string EF { get; set; }
        public string TXT_GATE_IN { get; set; }
        public string NAVIO_VIAGEM { get; set; }
        public string POD { get; set; }
        public string BRUTO { get; set; }
        public string LACRES_RESUMO { get; set; }
        public string IMO { get; set; }
        public int TAMANHO { get; set; }
        public string TIPO_BASICO { get; set; }
        public string YARD_ATUAL { get; set; }
        public string GMCI { get; set; }
        public string SIGVIG { get; set; }
        public string TEXTO_CONTEXTUAL { get; set; }
        public bool FLAG_REEFER_DESLIGADO_VISIVEL { get; set; }
        public bool FLAG_REEFER_LIGADO_VISIVEL { get; set; }
        public bool LBL_DTA_VISIVEL { get; set; }
        public bool LBL_DESOVA_VISIVEL { get; set; }
        public string LBL_DESOVA_TEXTO { get; set; }
        public string ALERTA_LINHA_9 { get; set; }
        public bool ALERTA_LINHA_9_VISIVEL { get; set; }
        public int FLAG_CNTR_BLOQ_CHAO_PATIO { get; set; }
        public string ISO { get; set; }
        public decimal TARA { get; set; }
        public decimal MAX_GROSS { get; set; }
        public string LACRE_IPA { get; set; }
        public string LACRE_IPA_TAG_ORIGEM { get; set; }
        public string LACRE1_GATE { get; set; }
        public string LACRE1_TAG { get; set; }
        public string LACRE2_GATE { get; set; }
        public string LACRE2_TAG { get; set; }
        public string LACRE3_GATE { get; set; }
        public string LACRE3_TAG { get; set; }
        public string LACRE_LONA { get; set; }
        public string TEMP_SETPOINT { get; set; }
        public string TEMP_MAX { get; set; }
        public string TEMP_MIN { get; set; }
        public bool FRAME_TRANSPORTE_INTERNO { get; set; }
        public List<CntrPatioBlDto> LOTES_BL { get; set; } = new List<CntrPatioBlDto>();
        public bool AVISO_REEFER_POSICAO_BLOQUEADO { get; set; }
        public string AVISO_REEFER_POSICAO_MSG { get; set; }
    }

    public class CntrPatioBlDto
    {
        public long LOTE { get; set; }
        public string BL { get; set; }
    }

    public class ProximoMovimentoDto
    {
        public string MOTIVO { get; set; }
        public string DT_PREV { get; set; }
    }

    public class VizinhoPilhaDto
    {
        public string ID_CONTEINER { get; set; }
        public string YARD { get; set; }
        public string DT_PREV { get; set; }
        public string MOTIVO { get; set; }
        public string PILHA { get; set; }
    }

    public class MovimentacaoCntrConfirmRequest
    {
        public long AUTONUM_CNTR { get; set; }
        public string ID_CONTEINER { get; set; }
        public string SISTEMA { get; set; }
        public string YARD_ORIGEM { get; set; }
        public string YARD_DESTINO { get; set; }
        public int MOTIVO_ID { get; set; }
        public int? CAMERA_ID { get; set; }
        public int TAMANHO { get; set; }
        public string EF { get; set; }
        public string TIPO_BASICO { get; set; }
        public decimal TARA { get; set; }
        public string GMCI { get; set; }
        public string LACRE_IPA { get; set; }
        public string LACRE1_GATE { get; set; }
        public string LACRE2_GATE { get; set; }
        public string LACRE3_GATE { get; set; }
        public string LACRE_LONA { get; set; }
        public string LACRE_IPA_TAG_ORIGEM { get; set; }
        public string LACRE1_TAG { get; set; }
        public string LACRE2_TAG { get; set; }
        public string LACRE3_TAG { get; set; }
        public bool LACRES_CONFERIDOS { get; set; }
        public int FLAG_CNTR_BLOQ_CHAO_PATIO { get; set; }
        public string IMO { get; set; }
        public bool CONFIRMOU_SAIDA { get; set; }
        public bool CONFIRMOU_VEICULO_SERV { get; set; }
    }

    public class MovimentacaoCntrConfirmResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public string CODIGO { get; set; }
        public bool EXIGE_CAMERA { get; set; }
        public bool EXIGE_CONFIRMACAO_SAIDA { get; set; }
        public bool EXIGE_VEICULO_SERV { get; set; }
    }

    public class VizinhancaPilhaRequest
    {
        public string ID_CONTEINER { get; set; }
        public string YARD_ATUAL { get; set; }
        public string PREFIXO_DESTINO { get; set; }
        public bool MODO_DESTINO { get; set; }
    }

    public class HistoricoShiftingRowDto
    {
        public string DATA { get; set; }
        public string ORIGEM { get; set; }
        public string DESTINO { get; set; }
        public string NMUSUARIO { get; set; }
        public string IDENTIFICACAO { get; set; }
        public string DESC_MOTIVO { get; set; }
    }

    public class AvariaCntrGridDto
    {
        public string LOCAL { get; set; }
        public string TIPO { get; set; }
        public long AUTONUM { get; set; }
        public bool DESTACAR_OPERADOR { get; set; }
    }

    public class AvariaHistDto
    {
        public string DT_AVARIA { get; set; }
        public string AVARIA { get; set; }
        public string USUARIO { get; set; }
    }

    public class AvariaIncluirRequest
    {
        public long AUTONUM_CNTR { get; set; }
        public string LOCAL_DESCR { get; set; }
        public List<string> TIPOS_DESCR { get; set; }
    }

    public class AvariaExcluirRequest
    {
        public long AUTONUM_CNTR { get; set; }
        public long AUTONUM_AVARIA { get; set; }
    }

    public class AvariaFinalizarRequest
    {
        public long AUTONUM_CNTR { get; set; }
        public string ID_CONTEINER { get; set; }
    }

    public class OpcaoListaDto
    {
        public string DESCR { get; set; }
    }
}
