using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models.ViewModels
{
    public class PosicionamentoViewModel
    {
        public string AUTONUM_CS { get; set; }
        public int ITEM { get; set; }
        public string QUANTIDADE { get; set; }

        public string POSICAO { get; set; }

        public string ARMAZEM { get; set; }
        public int AUTONUM_ROMANEIO { get; set; }
        public int AUTONUM_CNTR { get; set; }
        public int LOTE { get; set; }
        public string ID_CONTEINER { get; set; }

        public string PATIO { get; set; }

         public List<Armazens> LISTA_ARMAZENS { get; set; }

        public List<Posicionamento> LISTA_POSICIONAMENTOS { get; set; }
    }
}