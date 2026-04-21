using System.Collections.Generic;
using Romaneio.Classes;

namespace Romaneio.Models
{
    public class FotosRomaneio
    {
        public int AUTONUM_ROMANEIO { get; set; }
        public int ID_FOTO { get; set; }
        public int ID_TIPO_FOTO { get; set; }
        public int ID_TIPO_PROCESSO { get; set; }
        public string DESCR_TIPO { get; set; }
        public string DESCRICAO { get; set; }
        public string OBSERVACAO { get; set; }
        public string FOTO { get; set; }

        public List<TipoFotoDTO> LISTA_TIPO_FOTO { get; set; }

    }
}