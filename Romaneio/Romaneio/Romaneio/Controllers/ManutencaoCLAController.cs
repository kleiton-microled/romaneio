using Romaneio.Interfaces;
using Romaneio.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class ManutencaoCLAController : Controller
    {
        private readonly IManutencaoCLARepositorio _ManutencaoCLARepositorio;
        private readonly IDadosRepositorio _DadosRepositorio;

        public ManutencaoCLAController(IDadosRepositorio dadossrepositorio, IManutencaoCLARepositorio clarepositorio)
        {
            _ManutencaoCLARepositorio = clarepositorio;
            _DadosRepositorio = dadossrepositorio;

        }


        //MANUTENCAO CLA CARGA SOLTA
        public ActionResult Index(int LOTE = 0, int ITEM = 0, int AUTONUM_BREAK_BULK = 0)
        {
            LoteViewModel view = new LoteViewModel();

            if (LOTE != 0)
            {
                view.LISTA_ITENS = _DadosRepositorio.ListaItens(LOTE, 0).ToList();
            }
            if (ITEM != 0)
            {
                view.LISTA_CLA = _ManutencaoCLARepositorio.ListarCLA(LOTE, ITEM).ToList();

                view.AUTONUM_CS = _ManutencaoCLARepositorio.ObterAutonumCS(LOTE, ITEM);

            }
            if (AUTONUM_BREAK_BULK != 0)
            {
                var cla = _ManutencaoCLARepositorio.ObterCLA(AUTONUM_BREAK_BULK);
                if (cla != null)
                {
                    view.AUTONUM_BREAK_BULK = cla.AUTONUM_BREAK_BULK;
                    view.AUTONUM_CS = cla.AUTONUM_CS;
                    view.COMPRIMENTO = cla.COMPRIMENTO;
                    view.LARGURA = cla.LARGURA;
                    view.ALTURA = cla.ALTURA;
                    view.QUANTIDADE_CLA = cla.QUANTIDADE_CLA;
                    view.PESO = cla.PESO;

                }
            }

            return View(view);
        }

        [HttpPost]
        public JsonResult ExcluirCLA(int AUTONUM_BREAK_BULK)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }
            else
            {
                if (AUTONUM_BREAK_BULK != null)
                {
                    _ManutencaoCLARepositorio.ExcluirCLA(AUTONUM_BREAK_BULK);
                    return Json(new { success = true });
                }
                return Json(new { success = false, message = "Item não encontrado" });
            }
        }




        //SALVAR MANUTENCAO CLA CARGA SOLTA E ROMANEIO

        [HttpPost]
        public ActionResult SalvarDados(LoteViewModel DadosEntrada, string acao)
        {
            try
            {
                if (acao == "Salvar")
                {
                    if (DadosEntrada.LOTE == null)
                    {
                        TempData["MensagemErro"] = "Erro ao realizar salvar informações: informe o Lote!";
                        return RedirectToAction(nameof(Index), new { LOTE = DadosEntrada.LOTE, ITEM = DadosEntrada.ITEM });
                    }
                    if (DadosEntrada.ITEM == null)
                    {
                        TempData["MensagemErro"] = "Erro ao realizar salvar informações: informe o item!";
                        return RedirectToAction(nameof(Index), new { LOTE = DadosEntrada.LOTE, ITEM = DadosEntrada.ITEM });
                    }
                    else
                    {
                        if (_ManutencaoCLARepositorio.VerificarCubagem(DadosEntrada) == 1)
                        {
                            TempData["MensagemErro"] = "CLA resulta em cubagem superior/inferior ao permitido!";
                            // return RedirectToAction(nameof(Index), new { LOTE = DadosEntrada.LOTE, ITEM = DadosEntrada.ITEM });
                        }

                        string result = _ManutencaoCLARepositorio.SalvarDadosCLA(DadosEntrada);


                        if (result == "Informações salvas com sucesso!")
                        {
                            TempData["MensagemSucesso"] = "Informações salvas com sucesso!";
                        }
                        else
                        {
                            TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + result;
                        }

                        return RedirectToAction(nameof(Index), new { LOTE = DadosEntrada.LOTE, ITEM = DadosEntrada.ITEM });

                    }
                }
                else 
                {
                    if (DadosEntrada.AUTONUM_ROMANEIO == null)
                    {
                        TempData["MensagemErro"] = "Erro ao realizar salvar informações: necessario salvar romaneio!";
                        return RedirectToAction(nameof(Index), new { ROMANEIO = DadosEntrada.AUTONUM_ROMANEIO });
                    }
                    else
                    {

                        string result = _ManutencaoCLARepositorio.SalvarDadosCLARomaneio(DadosEntrada);

                        if (result == "Informações salvas com sucesso!")
                        {
                            TempData["MensagemSucesso"] = "Informações salvas com sucesso!";
                        }
                        else
                        {
                            TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + result;
                        }

                    }

                      return RedirectToAction(nameof(ConsultaCLARomaneio), new { AUTONUM_ROMANEIO = DadosEntrada.AUTONUM_ROMANEIO });

                }


            }
            catch (Exception erro)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + erro;
                return RedirectToAction(nameof(Index), new { LOTE = DadosEntrada.LOTE, ITEM = DadosEntrada.ITEM });
            }
        }




        //MANUTENCAO CLA ROMANEIO

        public ActionResult ConsultaCLARomaneio(int AUTONUM_ROMANEIO = 0, int AUTONUM_BREAK_BULK = 0)
        {
            LoteViewModel view = new LoteViewModel();

            view.AUTONUM_ROMANEIO = AUTONUM_ROMANEIO;
            view.LISTA_CLA = _ManutencaoCLARepositorio.ListarCLARomaneio(AUTONUM_ROMANEIO).ToList();

            if (AUTONUM_BREAK_BULK != 0)
            {
                var cla = _ManutencaoCLARepositorio.ObterCLARomaneio(AUTONUM_BREAK_BULK);
                if (cla != null)
                {
                    view.AUTONUM_BREAK_BULK = cla.AUTONUM_BREAK_BULK;
                    view.AUTONUM_CS = cla.AUTONUM_CS;
                    view.COMPRIMENTO = cla.COMPRIMENTO;
                    view.LARGURA = cla.LARGURA;
                    view.ALTURA = cla.ALTURA;
                    view.QUANTIDADE_CLA = cla.QUANTIDADE_CLA;
                    view.PESO = cla.PESO;

                }
            }

            return View("Index", view);
        }

        [HttpPost]
        public JsonResult ExcluirCLARomaneio(int AUTONUM_BREAK_BULK)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }
            else
            {
                if (AUTONUM_BREAK_BULK != null)
                {
                    _ManutencaoCLARepositorio.ExcluirCLARomaneio(AUTONUM_BREAK_BULK);
                    return Json(new { success = true });
                }
                return Json(new { success = false, message = "Item não encontrado" });
            }
        }



    }
}