using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Classes;
using Romaneio.Models;


namespace Romaneio.Interfaces
{
    public interface IFechamentosRepositorio
    {
        IEnumerable<Lotes> ListaLotesFechados(int cntr);
        IEnumerable<Conteiners> ConsultarFechamentos(int PATIO);
        LotesDTO ObterDadosRomaneioFechado(int cntr = 0, int lote = 0, int item = 0, int AUTONUM_ROMANEIO = 0);
        IEnumerable<LotesDTO> ConsultarFechamentosLotesConteiners(string cntr);
        LotesDTO ObterDadosLote(int cntr = 0, int lote = 0, int item = 0);
    }
}