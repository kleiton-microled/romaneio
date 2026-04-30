using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Linq;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class AssociacaoMarcantesController : DefaultController
    {
        private readonly IAssociacaoMarcantesRepositorio _associacaoMarcantesRepositorio;

        public AssociacaoMarcantesController(IAssociacaoMarcantesRepositorio associacaoMarcantesRepositorio)
        {
            _associacaoMarcantesRepositorio = associacaoMarcantesRepositorio;
        }

        [HttpGet]
        public ActionResult Index()
        {
            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }

            var view = new AssociacaoMarcantesViewModel();
            return View(view);
        }

        [HttpPost]
        public JsonResult ObterConteinersPorLote(int? lote)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var conteineres = _associacaoMarcantesRepositorio.ObterConteinersPorLote(lote);
            return Json(new { success = true, conteineres });
        }

        [HttpPost]
        public JsonResult Filtrar(int? lote, string cntr)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var resultado = _associacaoMarcantesRepositorio.Filtrar(lote, cntr);
            var loteNaoEncontrado = lote.HasValue &&
                                    (resultado.ITENS == null || !resultado.ITENS.Cast<object>().Any());
            return Json(new
            {
                success = true,
                modoExp = resultado.MODO_EXP,
                itens = resultado.ITENS,
                loteNaoEncontrado = loteNaoEncontrado
            });
        }

        [HttpPost]
        public JsonResult CarregarTemporarios(bool modoExp, int autonumRef)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var itens = _associacaoMarcantesRepositorio.CarregarTemporarios(modoExp, autonumRef);
            return Json(new { success = true, itens });
        }

        [HttpPost]
        public JsonResult AdicionarTemp(AssociacaoMarcantesAdicionarRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var mensagem = _associacaoMarcantesRepositorio.AdicionarTemp(request);
            if (!string.IsNullOrWhiteSpace(mensagem))
            {
                return Json(new { success = false, message = mensagem });
            }

            var itens = _associacaoMarcantesRepositorio.CarregarTemporarios(request.MODO_EXP, request.AUTONUM_REF);
            return Json(new { success = true, itens });
        }

        [HttpPost]
        public JsonResult RemoverTemp(AssociacaoMarcantesRemoverRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var mensagem = _associacaoMarcantesRepositorio.RemoverTemp(request);
            if (!string.IsNullOrWhiteSpace(mensagem))
            {
                return Json(new { success = false, message = mensagem });
            }

            var itens = _associacaoMarcantesRepositorio.CarregarTemporarios(request.MODO_EXP, request.AUTONUM_REF);
            return Json(new { success = true, itens });
        }

        [HttpPost]
        public JsonResult Salvar(AssociacaoMarcantesSalvarRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var mensagem = _associacaoMarcantesRepositorio.Salvar(request);
            if (!string.Equals(mensagem, "MARCANTES ASSOCIADOS COM SUCESSO", StringComparison.OrdinalIgnoreCase))
            {
                return Json(new { success = false, message = mensagem });
            }

            return Json(new { success = true, message = mensagem });
        }
    }
}
