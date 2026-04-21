using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Classes;
using Romaneio.Models;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface ICompartilhaScannerRepositorio
    {
        IEnumerable<Destinos> ConsultarDestinos();

        IEnumerable<AutorizacaoSaida> ConsultarDados(int patio, string viagem, string cntr, string status);

        string Autorizar( CompartilhaScannerViewModel dados);

    }
}