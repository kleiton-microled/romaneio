using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Classes;
using Romaneio.Models;


namespace Romaneio.Interfaces
{
    public interface IRomaneioRepositorio
    {
        //ROMANEIO
        IEnumerable<LotesDTO> ConsultarLotesConteiners(string cntr);
        IEnumerable<Conteiners> ConsultarConteiners(int PATIO);
        IEnumerable<Lotes> ListaLotes(int cntr);
        LotesDTO ObterDadosLote(int cntr = 0, int lote = 0, int item = 0);
        LotesDTO ObterDadosRomaneio(int cntr = 0, int lote = 0, int item = 0,int AUTONUM_ROMANEIO = 0);
        LotesDTO ObterDadosHubRomaneio(int lote);
        string SalvarDados(LoteViewModel dados);
        string SalvarDadosHUBRomaneio(LoteViewModel dados);
        string Finalizar(int romaneio);
        int VerificarCubagem(int romaneio);


        ////ACRESCIMO
        string InserirLote(int cntr);


       


        //FOTOS
        IEnumerable<TipoFotoDTO> ListaTipoFoto();
        string InsereNovaFoto(FotosRomaneio dados);

        IEnumerable<FotosRomaneio> ConsultarFotosRomaneio(int AUTONUM_ROMANEIO);

        string BuscarFoto(int ID_FOTO);

        string ExcluirFoto(int ID_FOTO, int AUTONUM_ROMANEIO);



        //AVARIAS
        IEnumerable<AvariasRomaneio> ConsultarAvariasRomaneio(int AUTONUM_ROMANEIO);
       
        string ExcluirAvaria(int AUTONUM_ROMANEIO_AVARIAS, int AUTONUM_ROMANEIO);

        string InsereNovaAvaria(AvariasRomaneio dados);

        int ContaAvarias(int romaneio);

        string ExcluirTodaAvaria(int AUTONUM_ROMANEIO);



    }
}