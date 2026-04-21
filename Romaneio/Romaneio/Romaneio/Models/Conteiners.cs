using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class Conteiners
    {
        public int AUTONUM_CNTR { get; set; }
        public int LOTE { get; set; }
        public string ID_CONTEINER { get; set; }
        public string TAMANHO { get; set; }
        public string DESCR_TIPO { get; set; }
        public DateTime DT_ENTRADA { get; set; }
        public DateTime DT_AGENDAMENTO { get; set; }
        public string LACRE_ORIGEM { get; set; }
        public string LACRE_IPA { get; set; }
        public string VIAGEM { get; set; }
        public string BOL_DESUNIT { get; set; }
        public string DDA { get; set; }
        public string SOLICITADO { get; set; }
        public string ABERTURA { get; set; }
        public string FECHAMENTO { get; set; }
        public string DESCRICAO_SIGVIG { get; set; }
        public string POSICAO { get; set; }
    }
}