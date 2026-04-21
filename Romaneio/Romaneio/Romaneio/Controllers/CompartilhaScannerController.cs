using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class CompartilhaScannerController : Controller
    {
        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly ICompartilhaScannerRepositorio _CompartilhaScannerRepositorio;

        public CompartilhaScannerController(IDadosRepositorio dadossrepositorio, ICompartilhaScannerRepositorio compartilhaScannerRepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _CompartilhaScannerRepositorio = compartilhaScannerRepositorio;
        }
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult AutorizacaoSaida(int autonum = 0, int patio = 0, string viagem = "", string cntr = "", string status = "")
        {
             var model = new CompartilhaScannerViewModel
             {
                AUTONUM = autonum,
                PATIO = patio,
                VIAGEM = viagem,
                ID_CONTEINER = cntr,
                STATUS = status,
                LISTA_PATIO = _DadosRepositorio.ConsultarPatios().ToList(),
                LISTA_DESTINO = _CompartilhaScannerRepositorio.ConsultarDestinos().ToList(),
                LISTA_DADOS = _CompartilhaScannerRepositorio.ConsultarDados(patio, viagem, cntr, status).ToList(),
             };

            return View("AutorizacaoSaida", model);
        }

        [HttpPost]
        public ActionResult Autorizar(CompartilhaScannerViewModel dados)
        {
            if (dados.AUTONUM == 0)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Tipo obrigatorio!";
                return RedirectToAction(nameof(AutorizacaoSaida), new { autonum = dados.AUTONUM, patio = dados.PATIO, viagem = dados.VIAGEM, status = dados.STATUS });
            }
            if (dados.DESTINO == 0)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Local obrigatorio!";
                return RedirectToAction(nameof(AutorizacaoSaida), new { autonum = dados.AUTONUM, patio = dados.PATIO, viagem = dados.VIAGEM, status = dados.STATUS });
            }
            if (dados.SISTEMA == "")
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Local obrigatorio!";
                return RedirectToAction(nameof(AutorizacaoSaida), new { autonum = dados.AUTONUM, patio = dados.PATIO, viagem = dados.VIAGEM, status = dados.STATUS });
            }
            if (dados.ID_CONTEINER == "")
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Local obrigatorio!";
                return RedirectToAction(nameof(AutorizacaoSaida), new { autonum = dados.AUTONUM, patio = dados.PATIO, viagem = dados.VIAGEM, status = dados.STATUS });
            }

            _CompartilhaScannerRepositorio.Autorizar(dados);
            TempData["MensagemSucesso"] = "Salvo com sucesso!";

            return RedirectToAction(nameof(AutorizacaoSaida), new { autonum = dados.AUTONUM, patio = dados.PATIO, viagem = dados.VIAGEM, status = dados.STATUS });
        }

    }
}