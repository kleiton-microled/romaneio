using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IInventarioCargaSoltaArmazemRepositorio
    {
        List<InventarioCsComboDto> ListarArmazens(int patioUsuario);
        List<InventarioCsComboDto> ListarMotivos();
        InventarioCsMarcanteEstadoResponse ObterEstadoMarcante(int marcante, int patioUsuario);
        InventarioCsMarcanteEstadoResponse ObterEstadoLote(long loteBl, int marcante, long autonumCsYard);
        InventarioCsItemDetalheResponse ObterDetalheItem(string idGravacao);
        List<InventarioCsHistoricoLinhaDto> ListarHistorico(int marcante);
        InventarioCsResolverPrateleiraResponse ResolverPrateleira(string texto, int patioUsuario);
        InventarioCsSalvarResult Salvar(InventarioCsSalvarRequest request, int codUsuario, int patioUsuario);
    }
}
