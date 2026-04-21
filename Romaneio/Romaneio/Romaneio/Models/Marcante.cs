using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class Marcante
    {
        public int AUTONUM { get; set; }
        public int AUTONUM_CARGA { get; set; }
        public int AUTONUM_CARGA_AC { get; set; }
        public int VOLUMES { get; set; }
        public int AUTONUM_CS_YARD { get; set; }
        public int AUTONUM_CS_YARD_AC { get; set; }
        public int FLAG_RECONHECIMENTO { get; set; }
        public int LOTE_IMPRESSO { get; set; }
        public string ID_CONTEINER_IMPRESSO { get; set; }
        public int ARMAZEM_REC { get; set; }
        public string YARD_REC { get; set; }
        public int AUTONUM_CEXP { get; set; }
        public int PATIO { get; set; }



        public int LOTE { get; set; }
        public string BL { get; set; }
        public string EMBALAGEM { get; set; }
        public string IMPORTADOR { get; set; }
        public string NAVIO { get; set; }
        public int QTDE { get; set; }
        public int AUTONUMCS { get; set; }
        public string MERCADORIA { get; set; }
        public string DATA_ENTRADA { get; set; }
        public string TIPO_DOC { get; set; }
        public string DESCR_ARMAZEM { get; set; }
        public string POSICAO { get; set; }
        public string ID_GRAVACAO { get; set; }

    }
}