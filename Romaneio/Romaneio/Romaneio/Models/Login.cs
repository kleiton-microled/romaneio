using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class Login
    {
        public int AUTONUM { get; set; }
        public string NOME { get; set; }
        public string LOGIN { get; set; }
        public string SENHA { get; set; }
        public int FLAG_CARREGAMENTO_MARCANTE_BIP { get; set; }
    }
}