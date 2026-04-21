using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class Lotes
    {
        public int AUTONUM { get; set; }
        public int LOTE { get; set; }
        public string BL { get; set; }
        public string ITEM { get; set; }
        public string MARCA { get; set; }
        public string QUANTIDADE { get; set; }
        public string NCM { get; set; }
        public string DESCRICAO { get; set; }
        public string LOCALIZACAO { get; set; }
        public string IMPORTADOR { get; set; }
        public string SHIPPER { get; set; }
    }
}