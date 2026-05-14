using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IManutencaoMarcantesArmazemRepositorio
    {
        ManutencaoMarcantesArmazemFiltrarResult Filtrar(int lote);

        ManutencaoMarcantesArmazemFiltrarResult Salvar(ManutencaoMarcantesArmazemSalvarRequest request);
    }
}
