using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Classes;
using Romaneio.Models;


namespace Romaneio.Interfaces
{
    public interface IMovimentacaoCSRepositorio
    {
         Marcante ConsultarMarcante(string marcante);
        Marcante ConsultarMarcanteInventArmazem(string marcante);
        string ConsultarArmazem (int armazem_rec);
        MovimentacaoCSViewModel ConsultarCS(int autonum_cs);
        MovimentacaoCSViewModel ConsultarLote(int lote, string marcante);
        IEnumerable<Itens> ConsultarItens(int lote, string marcante, int autonum_cs_yard);
        IEnumerable<Motivo> ConsultarMotivo();
        string SalvarDados(MovimentacaoCSViewModel dados);

        MovimentacaoCSViewModel CarregaDadosItem(string id_Gravacao);
        IEnumerable<Historico> Carrega_Hist(string marcante);

        IEnumerable<CargasCT> CarregaCargasCT(string armazem);
        IEnumerable<Lacres> CarregaLacresCT(int autonum_armazens_ipa);
 
        string InsereLacreCT(int AUTONUM_ARMAZENS_IPA, string LACRE);
        string ExcluirLacreCT(int AUTONUM_LACRE);
    }
}