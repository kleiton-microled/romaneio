using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class AvariasRomaneio
    {
        public int AUTONUM_ROMANEIO { get; set; }
        public int AUTONUM_ROMANEIO_AVARIAS { get; set; }
       
        public string TIPO { get; set; }
        public string DESCR_TIPO { get; set; }
     
        public int LOCAL { get; set; }
        public string DESCR_LOCAL { get; set; }
       
        public int PROVIDENCIA { get; set; }
        public string DESCR_PROVIDENCIA { get; set; }
      
        public int COMPLEMENTO { get; set; }
        public string DESCR_COMPLEMENTO { get; set; }
        
        public string QTD { get; set; }
        public string PESO { get; set; }
        public string COMPLEMENTO_IPA { get; set; }
        public string DT_AVARIA { get; set; }

    }
}