using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Romaneio.Models
{
    public class ConsultaViewModel
    {
        public int Patio { get; set; }
        public List<Conteiners> Itens { get; set; }
        public int PaginaAtual { get; set; }
        public int TotalPaginas { get; set; }
        public int TamanhoPagina { get; set; }
    }
}