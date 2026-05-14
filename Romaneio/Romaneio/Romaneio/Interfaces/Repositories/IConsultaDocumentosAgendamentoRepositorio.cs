using System.Collections.Generic;
using Romaneio.Models.ViewModels;

namespace Romaneio.Interfaces
{
    public interface IConsultaDocumentosAgendamentoRepositorio
    {
        int? ResolverIdSolicitacaoLtlPorLote(int lote);

        List<DocumentoAgendamentoListaDto> ListarDocumentosCargaSolta(int autonumAgendamento, int lote);

        List<DocumentoAgendamentoListaDto> ListarDocumentosLtl(int idSolicitacaoLtl, int lote);

        DocumentoAgendamentoArquivoMetaDto ObterMetaArquivo(int autonumDocumento, int autonumAgendamento, bool ltl, int lote);

        byte[] ObterBlobDocumento(int autonumDocumento);

        void InserirBloqueioDocumento(int usuarioId, int autonumAgendamento);

        void RemoverBloqueioDocumento(int usuarioId, int autonumAgendamento);
    }
}
