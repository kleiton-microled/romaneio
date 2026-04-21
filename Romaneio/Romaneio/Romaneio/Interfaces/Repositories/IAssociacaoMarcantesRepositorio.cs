using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IAssociacaoMarcantesRepositorio
    {
        List<string> ObterConteinersPorLote(int? lote);
        AssociacaoMarcantesFiltroResult Filtrar(int? lote, string cntr);
        List<AssociacaoMarcanteTempDto> CarregarTemporarios(bool modoExp, int autonumRef);
        string AdicionarTemp(AssociacaoMarcantesAdicionarRequest request);
        string RemoverTemp(AssociacaoMarcantesRemoverRequest request);
        string Salvar(AssociacaoMarcantesSalvarRequest request);
    }
}
