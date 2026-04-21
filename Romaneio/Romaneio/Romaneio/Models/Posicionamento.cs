using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class Posicionamento
    {
        public int AUTONUM_CS { get; set; }
        public int AUTONUM_YARD { get; set; }

        public int ITEM { get; set; }
        public string QUANTIDADE { get; set; }

        public string POSICAO { get; set; }

        public string ARMAZEM { get; set; }
    }
}