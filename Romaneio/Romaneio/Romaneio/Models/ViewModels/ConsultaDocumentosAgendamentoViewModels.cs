using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class ConsultaDocumentosAgendamentoPageViewModel
    {
        public bool LTL { get; set; }
        public int LOTE { get; set; }
        public int AUTONUM_AGENDAMENTO { get; set; }
        public int? ID_SOLICITACAO_LTL { get; set; }
        public string TITULO { get; set; }
        public string Erro { get; set; }
    }

    public class DocumentoAgendamentoListaDto
    {
        public string DESCRICAO { get; set; }
        public int? LOTE { get; set; }
        public string NOME_IMG { get; set; }
        public string DT_INCLUSAO { get; set; }
        public int AUTONUM_AV_IMAGEM { get; set; }
        public int? TIPO_DOCUMENTO { get; set; }
    }

    public class DocumentoAgendamentoArquivoMetaDto
    {
        public string EXTENSAO { get; set; }
        public bool TEM_BLOB { get; set; }
    }
}
