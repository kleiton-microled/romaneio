using Romaneio.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Classes
{
    public class ConteinersDTO
    {
        public int AUTONUM { get; set; }
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
        public int AUTONUM_ARMAZENS_IPA { get; set; }
        public string DESCR_ARMAZEM { get; set; }
        public int AUTONUM_TIPO_CT { get; set; }
        public int FLAG_CT { get; set; }
        public string DESCR_TIPO_CT { get; set; }
        public List<CargasCT> CARGA_CT { get; set; }
        public List<Lacres> LACRES { get; set; }


    }
}