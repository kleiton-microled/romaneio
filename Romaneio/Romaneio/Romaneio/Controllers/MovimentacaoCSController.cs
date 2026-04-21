using Romaneio.Classes;
using Romaneio.Interfaces;
using Romaneio.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class MovimentacaoCSController : DefaultController //Controller
    {

        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly IMovimentacaoCSRepositorio _MovimentacaoCSRepositorio;


        public MovimentacaoCSController(IDadosRepositorio dadossrepositorio, IMovimentacaoCSRepositorio movimentacaorepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _MovimentacaoCSRepositorio = movimentacaorepositorio;

        }
        public ActionResult Index( string MARCANTE = "")
        {
            MovimentacaoCSViewModel view = new MovimentacaoCSViewModel();
           
            string PATIO = Session["Patio"].ToString();
            view.TIPOS_EMBALAGENS = _DadosRepositorio.TiposEmbalagens().ToList();
            view.LISTA_MOTIVO = _MovimentacaoCSRepositorio.ConsultarMotivo().ToList();
            view.LISTA_ARMAZEM = _DadosRepositorio.ListaArmazens(PATIO).ToList();

            if (MARCANTE != "")
            {
                Marcante DadosMarcante = _MovimentacaoCSRepositorio.ConsultarMarcante(MARCANTE);

                if (DadosMarcante.AUTONUM_CEXP == 0 && DadosMarcante.AUTONUM_CARGA > 0)
                {

                    MovimentacaoCSViewModel Dados = _MovimentacaoCSRepositorio.ConsultarCS(DadosMarcante.AUTONUM_CARGA);

                    if (DadosMarcante != null)
                    {
                        view.BL = Dados.BL;
                        view.QTD_MARCANTE = DadosMarcante.VOLUMES;
                        view.QTD_EMBALAGEM = DadosMarcante.VOLUMES;
                        view.QTD_LOCAL = DadosMarcante.VOLUMES;
                        view.AUTONUM_CS_YARD = DadosMarcante.AUTONUM_CS_YARD;
                        view.ID_CONTEINER = DadosMarcante.ID_CONTEINER_IMPRESSO;
                        view.LOTE = DadosMarcante.LOTE_IMPRESSO;
                        view.LOCAL = _MovimentacaoCSRepositorio.ConsultarArmazem(DadosMarcante.ARMAZEM_REC);
                        view.ARMAZEM = DadosMarcante.ARMAZEM_REC;
                        view.LISTA_ITENS = _MovimentacaoCSRepositorio.ConsultarItens(DadosMarcante.LOTE_IMPRESSO, MARCANTE, DadosMarcante.AUTONUM_CS_YARD).ToList();

                        Dados = _MovimentacaoCSRepositorio.ConsultarLote(DadosMarcante.LOTE_IMPRESSO, MARCANTE);
                        if (Dados != null)
                        {
                            view.MERCADORIA = Dados.MERCADORIA;
                            view.MARCA = Dados.MARCA;
                            view.ENT_DESOVA = Dados.ENT_DESOVA;
                            view.ID_CONTEINER = Dados.ID_CONTEINER;
                            view.IMO = Dados.IMO;
                            view.MOV_AGEND = Dados.MOV_AGEND;
                            view.ITEM = Dados.ITEM;

                        }
                    }
                }
                else 
                {
                    //CARGA DE EXPORTACAO
                    DadosMarcante = _MovimentacaoCSRepositorio.ConsultarMarcanteInventArmazem(MARCANTE);
                    view.BL = DadosMarcante.BL;
                    view.LOTE = DadosMarcante.LOTE ;
                    view.MERCADORIA = DadosMarcante.MERCADORIA;
                    view.ENT_DESOVA = DadosMarcante.DATA_ENTRADA; 
                    view.IMPORTADOR = DadosMarcante.IMPORTADOR;
                    view.EMBALAGEM = DadosMarcante.EMBALAGEM;
 
                    view.QTD_MARCANTE = DadosMarcante.QTDE;
                    view.QTD_EMBALAGEM = DadosMarcante.QTDE;
                    view.QTD_LOCAL = DadosMarcante.QTDE;
                    view.AUTONUM_CS_YARD = DadosMarcante.AUTONUMCS;
                    view.LOCAL = DadosMarcante.DESCR_ARMAZEM + " - " + DadosMarcante.POSICAO;
                    view.ARMAZEM = 0;
                }

            }


            return View(view);
        }


        [HttpPost]
        public ActionResult SalvarDados(MovimentacaoCSViewModel DadosEntrada)
        {
            try
            {

                if (DadosEntrada.MARCANTE == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: informe o MARCANTE!";
                    return RedirectToAction(nameof(Index), new { MARCANTE = DadosEntrada.MARCANTE });
                }
                if (DadosEntrada.ITEM == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: informe o ITEM!";
                    return RedirectToAction(nameof(Index), new { MARCANTE = DadosEntrada.MARCANTE });
                }
                if (DadosEntrada.MOTIVO == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: informe o MOTIVO!";
                    return RedirectToAction(nameof(Index), new { MARCANTE = DadosEntrada.MARCANTE });
                }
                if (DadosEntrada.LOCAL == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: informe o LOCAL!";
                    return RedirectToAction(nameof(Index), new { MARCANTE = DadosEntrada.MARCANTE });
                }
                else
                {
                    string result = _MovimentacaoCSRepositorio.SalvarDados(DadosEntrada);


                    if (result == "Informações salvas com sucesso!")
                    {
                        TempData["MensagemSucesso"] = "Informações salvas com sucesso!";
                    }
                    else
                    {
                        TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + result;
                    }

                    return RedirectToAction(nameof(Index), new { MARCANTE = DadosEntrada.MARCANTE });
                }

            }
            catch (Exception erro)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + erro;
                return RedirectToAction(nameof(Index), new { MARCANTE = DadosEntrada.MARCANTE });
            }
        }

        public JsonResult CarregaDadosItem(string ITEM)
        {
            try
            {
                var query = _MovimentacaoCSRepositorio.CarregaDadosItem(ITEM);

 
                if (query == null)
                {
                    retornoJson.Mensagem = "Conteiner não encontrado";
                    retornoJson.objetoRetorno = null;
                    retornoJson.possuiDados = false;
                    retornoJson.statusRetorno = "500";

                }
                else
                {
                    retornoJson.Mensagem = "";
                    retornoJson.objetoRetorno = query;
                    retornoJson.possuiDados = true;
                    retornoJson.statusRetorno = "200";

                }

                return Json(retornoJson, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                retornoJson.Mensagem = "Os dados não foram carregados";
                retornoJson.objetoRetorno = null;
                retornoJson.statusRetorno = "500";
                retornoJson.possuiDados = false;

                return Json(retornoJson, JsonRequestBehavior.AllowGet);
            }
        }


        public ActionResult Historico(string Marcante)
        {

            IEnumerable<Historico> dados = _MovimentacaoCSRepositorio.Carrega_Hist(Marcante);

            return PartialView("Historico", dados);  

        }


        public ActionResult CargasCT(int autonumArmazensIpa, string cntrArmazem)
        {    
 
            ConteinersDTO view = new ConteinersDTO();
            view.CARGA_CT = _MovimentacaoCSRepositorio.CarregaCargasCT(cntrArmazem).ToList();
            view.DESCR_ARMAZEM = cntrArmazem;
            view.AUTONUM_ARMAZENS_IPA = autonumArmazensIpa;

            return View("CargasCntrCT", view);  

        }

        private IEnumerable<Lacres> ObterLacres(int AUTONUM_ARMAZENS_IPA)
        {
            return _MovimentacaoCSRepositorio.CarregaLacresCT(AUTONUM_ARMAZENS_IPA);
        }
        
        public ActionResult CarregaLacresCT(int AUTONUM_ARMAZENS_IPA)
        {
            var dados = ObterLacres(AUTONUM_ARMAZENS_IPA);

            ViewBag.AutonumArmazem = AUTONUM_ARMAZENS_IPA;
            return PartialView("Lacres", dados); 

        }

        [HttpPost]
        public ActionResult InsereLacreCT(int AUTONUM_ARMAZENS_IPA, string LACRE)
        {
            if (AUTONUM_ARMAZENS_IPA == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: armazem não encontrado!";
                return RedirectToAction(nameof(CarregaLacresCT), new { AUTONUM_ARMAZENS_IPA = AUTONUM_ARMAZENS_IPA });
            }
            if (LACRE == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Lacre obrigatorio!";
                return RedirectToAction(nameof(CarregaLacresCT), new { AUTONUM_ARMAZENS_IPA = AUTONUM_ARMAZENS_IPA });
            }

            _MovimentacaoCSRepositorio.InsereLacreCT(AUTONUM_ARMAZENS_IPA, LACRE);

            var dados = _MovimentacaoCSRepositorio.CarregaLacresCT(AUTONUM_ARMAZENS_IPA);

            ViewBag.AutonumArmazem = AUTONUM_ARMAZENS_IPA;

            return PartialView("Lacres", dados);
        }

        [HttpPost]
        public JsonResult ExcluirLacreCT(int AUTONUM_LACRE )
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }
            else
            {
                if (AUTONUM_LACRE != null)
                {
                    _MovimentacaoCSRepositorio.ExcluirLacreCT(AUTONUM_LACRE);
                    return Json(new { success = true });
                }
                return Json(new { success = false, message = "Item não encontrado" });
            }
        }

    }
}