using System.Collections.Generic;

namespace Romaneio.Models.ViewModels
{
    public class AssociacaoMarcantesViewModel
    {
        public int? LOTE { get; set; }
        public string CNTR { get; set; }
        public bool MODO_EXP { get; set; }
        public List<string> LISTA_CNTR { get; set; } = new List<string>();
        public List<AssociacaoMarcantesLinhaDto> ITENS { get; set; } = new List<AssociacaoMarcantesLinhaDto>();
        public List<AssociacaoMarcanteTempDto> MARCANTES_TEMP { get; set; } = new List<AssociacaoMarcanteTempDto>();
    }

    public class AssociacaoMarcantesLinhaDto
    {
        public int INDICE { get; set; }
        public int LOTE { get; set; }
        public int ITEM { get; set; }
        public int QUANTIDADE { get; set; }
        public string EMBALAGEM { get; set; }
        public string ID_CONTEINER { get; set; }
        public string MERCADORIA { get; set; }
        public int AUTONUMCS { get; set; }
        public int AUTONUMCNTR { get; set; }
    }

    public class AssociacaoMarcanteTempDto
    {
        public int MARCANTE { get; set; }
        public int QTDE { get; set; }
        public string MARCANTE_FORMATADO => MARCANTE.ToString("000000000");
    }

    public class AssociacaoMarcantesFiltroResult
    {
        public bool MODO_EXP { get; set; }
        public List<AssociacaoMarcantesLinhaDto> ITENS { get; set; } = new List<AssociacaoMarcantesLinhaDto>();
    }

    public class AssociacaoMarcantesAdicionarRequest
    {
        public bool MODO_EXP { get; set; }
        public int AUTONUM_REF { get; set; }
        public int MARCANTE { get; set; }
        public int QTDE { get; set; }
    }

    public class AssociacaoMarcantesRemoverRequest
    {
        public bool MODO_EXP { get; set; }
        public int AUTONUM_REF { get; set; }
        public int MARCANTE { get; set; }
    }

    public class AssociacaoMarcantesSalvarRequest
    {
        public bool MODO_EXP { get; set; }
        public AssociacaoMarcantesLinhaDto ITEM_SELECIONADO { get; set; }
    }
}
