using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class Lacres
    {
     public int AUTONUM { get; set; }
     public int AUTONUM_ARMAZENS_IPA { get; set; }
     public string LACRE { get; set; }
     public string DT_LANCAMENTO { get; set; }
     public string DT_INATIVACAO { get; set; }
     public int FLAG_ATIVO { get; set; }
    }
}