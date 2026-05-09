using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface ILocalizacaoCarregamentosRepositorio
    {
        string ObterDescricaoPatio(int patioUsuario);
        LocalizacaoCarregamentosFiltroResult Filtrar(LocalizacaoCarregamentosFiltroRequest request, int patioUsuario);
        MoverSaidaResult MoverParaSaida(MoverSaidaRequest request, int patioUsuario, int usuarioId);
    }
}
