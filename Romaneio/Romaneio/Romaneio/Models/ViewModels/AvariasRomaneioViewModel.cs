using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models.ViewModels
{
    public class AvariasRomaneioViewModel
    {
        public int AUTONUM_ROMANEIO { get; set; }
        public int PATIO { get; set; }
        public int LOTE { get; set; }
        public int ID_CONTEINER { get; set; }
        public int ITEM { get; set; }
        public int AUTONUM_CNTR { get; set; }
        public int AUTONUM_CS { get; set; }
        public int AUTONUM_ROMANEIO_AVARIAS { get; set; }   
        public string TIPO { get; set; }
        public List<Avarias> LISTA_TIPO { get; set; }
        public int LOCAL { get; set; }
        public List<Local> LISTA_LOCAL { get; set; }

        public int PROVIDENCIA { get; set; }

        public List<Providencia> LISTA_PROVIDENCIA { get; set; }

        public int COMPLEMENTO { get; set; }

        public List<Complemento> LISTA_COMPLEMENTO { get; set; }
        public string QTD { get; set; }
        public string PESO { get; set; }
        public string COMPLEMENTO_IPA { get; set; }

        public string DT_AVARIA { get; set; }
        public List<AvariasRomaneio> LISTA_AVARIAS { get; set; }

    }
}