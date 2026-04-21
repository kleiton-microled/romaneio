using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models;
using System.Configuration;
using System.IO;
using Romaneio.Models.ViewModels;
using Romaneio.Classes;
using Romaneio.Extensions;

namespace Romaneio.Controllers
{
    public class AvariasController : Controller
    {
        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly IRomaneioRepositorio _RomaneioRepositorio;

        public AvariasController(IDadosRepositorio dadossrepositorio, IRomaneioRepositorio romaneioRepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _RomaneioRepositorio = romaneioRepositorio;
        }


        [HttpGet]
        public ActionResult ConsultarAvarias(int AUTONUM_ROMANEIO = 0)
        {

            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                IEnumerable<AvariasRomaneio> ListaAvarias = _RomaneioRepositorio.ConsultarAvariasRomaneio(AUTONUM_ROMANEIO).ToList();
                LotesDTO DadosBasicos = _RomaneioRepositorio.ObterDadosRomaneio(0, 0, 0, AUTONUM_ROMANEIO);
                var model = new AvariasRomaneioViewModel
                {
                    AUTONUM_ROMANEIO = AUTONUM_ROMANEIO,
                    LISTA_AVARIAS = ListaAvarias.ToList(),
                    ITEM = DadosBasicos.ITEM.ToInt(),
                    PATIO = DadosBasicos.PATIO,
                    LOTE = DadosBasicos.LOTE.ToInt(),
                    AUTONUM_CNTR = DadosBasicos.AUTONUM_CNTR.ToInt(),
                    LISTA_TIPO = _DadosRepositorio.ListaTipoAvarias().ToList(),
                    LISTA_LOCAL = _DadosRepositorio.ListaLocalAvarias().ToList(),
                    LISTA_PROVIDENCIA = _DadosRepositorio.ListaProvidenciaAvarias().ToList(),
                    LISTA_COMPLEMENTO = _DadosRepositorio.ListaComplementoAvarias().ToList(),
                };

                return View("Avarias", model);
            }
        }


        [HttpPost]
        public ActionResult SalvarAvaria(AvariasRomaneio dados)
        {
            if (dados.TIPO == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Tipo obrigatorio!";
                return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
            }
            if (dados.LOCAL == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Local obrigatorio!";
                return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
            }
            if (dados.COMPLEMENTO == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Complemento obrigatorio!";
                return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
            }
            if (dados.PROVIDENCIA == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Providencia obrigatorio!";
                return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
            }
            if (dados.QTD == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Quantidade obrigatoria!";
                return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
            }
            if (dados.PESO == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Peso obrigatorio!";
                return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
            }
            _RomaneioRepositorio.InsereNovaAvaria(dados);

            return RedirectToAction(nameof(ConsultarAvarias), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
        }

        [HttpPost]
        public JsonResult ExcluirAvaria(int AUTONUM_ROMANEIO_AVARIAS, int AUTONUM_ROMANEIO)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }
            else
            {
                if (AUTONUM_ROMANEIO_AVARIAS != null)
                {
                    _RomaneioRepositorio.ExcluirAvaria(AUTONUM_ROMANEIO_AVARIAS, AUTONUM_ROMANEIO);
                    return Json(new { success = true });
                }
                return Json(new { success = false, message = "Item não encontrado" });
            }
        }
    }


}