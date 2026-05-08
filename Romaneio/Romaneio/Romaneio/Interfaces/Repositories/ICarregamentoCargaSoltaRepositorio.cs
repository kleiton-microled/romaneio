using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface ICarregamentoCargaSoltaRepositorio
    {
        List<VeiculoCarregamentoDto> ListarVeiculos(int patioUsuario);
        List<OrdemCarregamentoDto> ListarOrdens(string placaCavalo);
        List<MarcanteCarregadoDto> ListarCarregamentos(string placaCavalo, int autonumCs);
        List<LocalArmazemDto> ListarLocais(int autonumCs);
        MarcanteDadosDto BuscarMarcante(int marcante);
        CarregarItemResult Carregar(CarregarItemRequest request, int codUsuario);
        CancelarItemResult Cancelar(CancelarItemRequest request, int codUsuario);
        DocumentoOcResult ObterDocumento(int loteBl);
    }
}
