using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IMovimentacaoContainerPatioRepositorio
    {
        string ObterDescricaoPatio(int patioUsuario);
        MovimentacaoContainerPatioPermissoesDto ObterPermissoesAvarias(int usuarioId);
        List<MotivoMovimentacaoDto> ListarMotivosAtivos();
        List<CameraPatioDto> ListarCameras(int patioUsuario);
        bool MotivoExigeCamera(int motivoAutonum);
        int? ObterCameraPadraoPorYard(int patioUsuario, string yardDestino);

        CntrPatioBuscaFinalResult BuscarPorFinal(string final, int patioUsuario);
        CntrPatioDetalheDto CarregarDetalhe(string idConteiner, int patioUsuario, int usuarioId);
        List<ProximoMovimentoDto> ProximosMovimentos(string idConteiner);
        List<VizinhoPilhaDto> VizinhancaPilha(VizinhancaPilhaRequest request, int patioUsuario);

        MovimentacaoCntrConfirmResult ConfirmarMovimentacao(MovimentacaoCntrConfirmRequest request, int patioUsuario, int usuarioId);

        List<HistoricoShiftingRowDto> HistoricoShifting(string idConteiner, string sistema);
        List<AvariaHistDto> HistoricoAvarias(long autonumCntr);
        List<AvariaCntrGridDto> ListarAvariasAbertas(long autonumCntr);
        List<string> ListarDescricoesAvariasLocal();
        List<string> ListarDescricoesAvariasTipo();
        bool AvariasFinalizadas(long autonumCntr);
        void IncluirAvaria(long autonumCntr, string localDescr, string tipoDescr, int usuarioId);
        void ExcluirAvaria(long autonumAvaria);
        void FinalizarAvarias(long autonumCntr);
    }
}
