using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Classes
{
    public class LotesDTO
    {
        public int AUTONUM_ROMANEIO { get; set; }
        public int AUTONUM_CNTR { get; set; }
        public int LOTE { get; set; }
        public string BL { get; set; }
        public string ID_CONTEINER { get; set; }

        public string INICIO { get; set; }

        public string DEMURRAGE { get; set; }

        public string TERMINO { get; set; }

        public string CONFERENTE { get; set; }

        public string SOLICITANTE { get; set; }

        public string FISCAL { get; set; }
        public string DESOVA { get; set; }
        public string INSPECAO_MAPA { get; set; }
        public List<int> AVARIAS_SELECIONADAS { get; set; }
        public string VOLUME { get; set; }

        public string PESO_BRUTO { get; set; }
        public string PESO_APURADO { get; set; }

        public string IMO { get; set; }

        public string UNDG { get; set; }
        public string GENERO { get; set; }

        public string ITEM { get; set; }
        public string MARCA { get; set; }
        public string QUANTIDADE { get; set; }

        public string EMBALAGEM { get; set; }

        public string MERCADORIA { get; set; }
        public string NCM { get; set; }
         public string LOCALIZACAO { get; set; }

        public string LOCAL { get; set; }

        public string SUB_CLASSE { get; set; }

        public string DESCRICAO { get; set; }

        public string IMPORTADOR { get; set; }
        public string SHIPPER { get; set; }

        public string OBS { get; set; }

        public bool FL_OUTRA_AVARIA { get; set; }
        public string DESCR_OUTRA_AVARIA { get; set; }
        public bool FL_FALTA { get; set; }
        public string DESCR_FALTA { get; set; }
        public bool FL_ACRESCIMO { get; set; }
        public string DESCR_ACRESCIMO { get; set; }
        public bool FL_AGUARDANDO_RECONHECIMENTO { get; set; }
        public string DESCR_AGUARDANDO_RECONHECIMENTO { get; set; }
        public string EMPRESA { get; set; }
        public int PATIO { get; set; }
        public bool FL_SEM_AVARIA { get; set; }
       
        //
        public bool FLAG_HUB { get; set; }
        public string HUB { get; set; }
        public string CIDADE_HUB { get; set; }
        public string QUANTIDADE_CLA { get; set; }
        public string CUBAGEM { get; set; }
        public decimal PESO { get; set; }
        public decimal COMPRIMENTO { get; set; }
        public decimal LARGURA { get; set; }
        public decimal ALTURA { get; set; }
        public int AUTONUM_CNTR_DESTINO { get; set; }
        public int AUTONUM_ROMANEIO_BREAK_BULK { get; set; }
        public int AUTONUM_BREAK_BULK { get; set; }
        public int AUTONUM_CS { get; set; }

    }
}