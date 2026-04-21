using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Classes;
using Romaneio.Models;


namespace Romaneio.Interfaces
{
    public interface IManutencaoCLARepositorio
    {
        //CARGA SOLTA
        IEnumerable<CLA> ListarCLA(int lote, int item);
        int ObterAutonumCS(int lote, int item);
        LoteViewModel ObterCLA(int AUTONUM_BREAK_BULK);
        string SalvarDadosCLA(LoteViewModel dados);
        int VerificarCubagem(LoteViewModel dados);
        string ExcluirCLA(int AUTONUM_BREAK_BULK);


        //ROMANEIO
        IEnumerable<CLA> ListarCLARomaneio(int AUTONUM_ROMANEIO);
        LoteViewModel ObterCLARomaneio(int AUTONUM_BREAK_BULK);
        string SalvarDadosCLARomaneio(LoteViewModel dados);
        string ExcluirCLARomaneio(int AUTONUM_BREAK_BULK);

    }
}