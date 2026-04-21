using Romaneio.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models
{
    public class MovimentacaoCSViewModel
    {
        public int AUTONUM_CS_YARD { get; set; }
        public int MARCANTE { get; set; }
        public int QTD_MARCANTE { get; set; }
        public int LOTE { get; set; }
        public string BL { get; set; }
        public string ITEM { get; set; }
        public List<Itens> LISTA_ITENS { get; set; } = new List<Itens>();
        public int QTD_EMBALAGEM { get; set; }
        public List<Embalagens> TIPOS_EMBALAGENS { get; set; } = new List<Embalagens>();
        public string EMBALAGEM { get; set; }
        public List<Armazens> LISTA_ARMAZEM { get; set; } = new List<Armazens>();
        public int ARMAZEM { get; set; }
        public string DESCR_ARMAZEM { get; set; }
        public string LOCAL { get; set; }       
        public string MERCADORIA { get; set; }
        public string MARCA { get; set; }
        public string IMPORTADOR { get; set; }
        public string ENT_DESOVA { get; set; }
        public int AUTONUM_CNTR { get; set; }
        public string ID_CONTEINER { get; set; }
        public string IMO { get; set; }
        public string MOV_AGEND { get; set; }
        public string ETQ_PRATELEIRA { get; set; }
        public int QTD_LOCAL { get; set; }
        public List<Motivo> LISTA_MOTIVO { get; set; } = new List<Motivo>();
        public int MOTIVO { get; set; }
        public bool FL_FRENTE { get; set; }
        public bool FL_FUNDO { get; set; }
        public bool FL_LD { get; set; }
        public bool FL_LE { get; set; }
 
    }

}