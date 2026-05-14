using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IContainerReeferPatioRepositorio
    {
        string ObterDescricaoPatio(int patioUsuario);

        CntrReeferBuscaFinalResult BuscarPorFinal(string final, int patioUsuario);

        ConteinerReeferDetalheDto CarregarConteiner(string idConteiner, int patioUsuario);

        IReadOnlyList<ReeferHistoricoLinhaDto> ListarHistorico(long autonumIpa, long autonumRdx, long autonumOp);

        SalvarMonitoramentoResult SalvarMonitoramento(SalvarMonitoramentoRequest request, int patioUsuario, int usuarioId);

        RegistrarPlugOffResult RegistrarPlugOff(long autonumIpa, int patioUsuario);

        IReadOnlyList<ReeferEntradaPrevistaDto> ListarEntradasPrevistas(int patioUsuario);

        IReadOnlyList<ReeferEstoqueLinhaDto> ListarEstoqueReefer(int patioUsuario, bool apenasDesligados, bool comAgendamentoSaida, bool posicionados);

        IReadOnlyList<ReeferDesligadoLinhaDto> ListarUnidadesDesligadas(int patioUsuario);
    }
}
