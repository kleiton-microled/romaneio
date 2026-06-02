using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Classes;
using Romaneio.Models;


namespace Romaneio.Interfaces
{
    public interface IDadosRepositorio
    {
        IEnumerable<TiposPatios> ConsultarPatios();
        Login ConsultarUsuario(string usuario);
        string ConsultarNomeUsuario(string id);
        IEnumerable<Solicitantes> TiposSolicitantes();
        IEnumerable<Embalagens> TiposEmbalagens();
        IEnumerable<TipoDesova> TiposDesova();
        IEnumerable<Generos> TiposGeneros();
        IEnumerable<RelacaoUnidadesViewModel> ConsultarUnidades(string lote);
        IEnumerable<NCMViewModel> ConsultarNCM(string lote);
        DescricaoDTO ConsultarDescr(string lote);
        LocalizacaoDTO ConsultarLocalizacao(string lote);
        IEnumerable<Posicionamento> ConsultarPosicoesItem(string patio, string cntr, string lote, string item);
        IEnumerable<ConteinersDTO> ConteinersDestino(int patio);
        IEnumerable<Armazens> ListaArmazens(string patio);
        IEnumerable<Armazens> ListaArmazensAbertos(string patio);

        IEnumerable<Itens> ListaItens(int lote, int cntr);
        string ConsultarSolicitanteCNTR(string AUTONUM_ROMANEIO, string CNTR);
        string ConsultarDesovaCNTR(string AUTONUM_ROMANEIO, string CNTR);
        string ObterAutonumCS(string patio, string cntr, string lote, string item);

        string InsereYard(string AUTONUM_CS, string QTD, string ARM, string YARD);

        string AlteraYard(int romaneio, string qtd, string arm, string yard);

        string DeletaYard(int AUTONUM, int AUTONUM_CS);

        IEnumerable<Avarias> ListaTipoAvarias();

        IEnumerable<Local> ListaLocalAvarias();

        IEnumerable<Providencia> ListaProvidenciaAvarias();

        IEnumerable<Complemento> ListaComplementoAvarias();

        IEnumerable<ConteinersDTO> ConsultarClasseCT(string armazem);

        LotesDTO ObterDadosHubLote(int lote);

    }
}