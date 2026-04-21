using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Models.ViewModels
{
    public class CompartilhaScannerViewModel
    {
        public int AUTONUM { get; set; }
        public string STATUS { get; set; }
        public string VIAGEM { get; set; }

        public string SISTEMA { get; set; }
        public string ID_CONTEINER { get; set; }
        public int AUTONUM_GATE { get; set; }
        public int PATIO { get; set; }
        public string DESCR_PATIO { get; set; }
        public List<TiposPatios> LISTA_PATIO { get; set; } = new List<TiposPatios>();

        public int DESTINO { get; set; }
        public string DESCR_DESTINO { get; set; } 
        public List<Destinos> LISTA_DESTINO { get; set; } = new List<Destinos>();

        public List<AutorizacaoSaida> LISTA_DADOS { get; set; } = new List<AutorizacaoSaida>();
 
    }

    public class AutorizacaoSaida
    {
        public int AUTONUM { get; set; }
        public string SISTEMA { get; set; }
        public string ID_CONTEINER { get; set; }
        public string TAMANHO { get; set; }
        public string EF { get; set; }
        public string VIAGEM { get; set; }
        public string TERMINAL { get; set; }
        public string DESTINO { get; set; }
        public string DT_SAIDA_SCANNER { get; set; }       

    }

    public class Destinos
    {
        public int AUTONUM { get; set; }
        public string DESCRICAO { get; set; }
    }

 }