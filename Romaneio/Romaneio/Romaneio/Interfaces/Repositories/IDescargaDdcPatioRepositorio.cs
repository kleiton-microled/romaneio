using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IDescargaDdcPatioRepositorio
    {
        string ObterDescricaoPatio(int patioUsuario);

        IReadOnlyList<PlacaDdcListItemDto> ListarPlacasDdc(int patioUsuario);

        GateLotesResultDto ResolverGateELotes(long autonumRegistroSaida, int patioUsuario);

        CarregarCabecalhoResultDto CarregarCabecalho(long autonumRegistroSaida, long autonumLote, int patioUsuario);

        IReadOnlyList<DescargaArmazemItemDto> ListarItens(long cntr, long gate, int patioUsuario);

        IReadOnlyList<EmbalagemDto> ListarEmbalagens();

        DescargaDdcPatioOperacaoResult SalvarItem(SalvarDescargaDdcRequest request, int patioUsuario, int usuarioId);

        DescargaDdcPatioOperacaoResult ExcluirItem(long autonumItem, int patioUsuario);
    }
}
