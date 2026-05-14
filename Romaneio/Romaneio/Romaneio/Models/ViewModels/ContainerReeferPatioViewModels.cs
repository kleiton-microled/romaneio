using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class ContainerReeferPatioViewModel
    {
        public int PATIO { get; set; }
        public string DESCR_PATIO { get; set; }
    }

    public class CntrReeferBuscaFinalResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public string ID_CONTEINER { get; set; }
    }

    public class ConteinerReeferDetalheDto
    {
        public string ID_CONTEINER { get; set; }
        public long AUTONUM { get; set; }
        public string SISTEMA { get; set; }
        public string DATA_ENT_TEMP { get; set; }
        public string TXT_SET_POINT { get; set; }
        public string TXT_UMIDADE_SET { get; set; }
        public string TXT_VENTILACAO_SET { get; set; }
        public string TXT_ESCALA { get; set; }
        public string TXT_MIN { get; set; }
        public string TXT_MAX { get; set; }
        public bool ALERTA_REEFER_DESLIGADO { get; set; }
        public string PLUG_OFF_TEXTO { get; set; }
        public long AUTONUM_IPA { get; set; }
        public long AUTONUM_RDX { get; set; }
        public long AUTONUM_OP { get; set; }
        public string ERRO { get; set; }
    }

    public class ReeferHistoricoLinhaDto
    {
        public string DT_MONITORING { get; set; }
        public string TEMP_SUP { get; set; }
        public string TEMP_RET { get; set; }
        public string UMIDADE { get; set; }
        public string VENTILACAO { get; set; }
        public string USUARIO { get; set; }
        public long AUTONUM { get; set; }
    }

    public class SalvarMonitoramentoRequest
    {
        public string ID_CONTEINER { get; set; }
        public long AUTONUM_IPA { get; set; }
        public long AUTONUM_RDX { get; set; }
        public long AUTONUM_OP { get; set; }
        public string TXT_SET_POINT { get; set; }
        public string TXT_UMIDADE_SET { get; set; }
        public string TXT_VENTILACAO_SET { get; set; }
        public string TXT_TEMP_SUP { get; set; }
        public string TXT_TEMP_RET { get; set; }
        public string TXT_UMIDADE { get; set; }
        public string TXT_VENTILACAO { get; set; }
        public string TXT_ESCALA { get; set; }
        public bool CONFIRMAR_DIVERGENCIA_TEMPERATURA { get; set; }
    }

    public class SalvarMonitoramentoResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
        public bool EXIGE_CONFIRMACAO_DIVERGENCIA_TEMPERATURA { get; set; }
    }

    public class RegistrarPlugOffResult
    {
        public bool SUCESSO { get; set; }
        public string MENSAGEM { get; set; }
    }

    public class ReeferEntradaPrevistaDto
    {
        public string ID_CONTEINER { get; set; }
        public string TAMANHO { get; set; }
        public string SETPOINT { get; set; }
        public string DT_ATRACACAO { get; set; }
        public string ARM_ATRACACAO { get; set; }
        public string ENTRADA_PREVISTA { get; set; }
    }

    public class ReeferEstoqueLinhaDto
    {
        public string ID_CONTEINER { get; set; }
        public string DT_ENTRADA { get; set; }
        public string YARD { get; set; }
        public string PERIODO_INICIAL { get; set; }
        public string SETPOINT { get; set; }
        public string TEMP_RET_M1 { get; set; }
        public string DT1 { get; set; }
        public string TEMP_RET_M2 { get; set; }
        public string DT2 { get; set; }
        public int TEM_PLUG_OFF { get; set; }
    }

    public class ReeferDesligadoLinhaDto : ReeferEstoqueLinhaDto
    {
        public string DT_PLUG_OFF_FMT { get; set; }
    }
}
