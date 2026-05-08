using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IConsultaLiberacaoCarregamentoRepositorio
    {
        List<TipoVeiculoDto> ObterTiposVeiculo();
        ConsultaLiberacaoResultadoDto Filtrar(ConsultaLiberacaoFiltroRequest request, int patioUsuario);
        List<string> CarregarProtocolosSimultaneos(string protocolo);
        RegistrarLiberacaoResultDto Registrar(RegistrarLiberacaoRequest request, int usuarioId, int patioUsuario);
    }
}
