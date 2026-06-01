using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class ConsultaLiberacaoCarregamentoController : DefaultController
    {
        private readonly IConsultaLiberacaoCarregamentoRepositorio _repositorio;

        public ConsultaLiberacaoCarregamentoController(IConsultaLiberacaoCarregamentoRepositorio repositorio)
        {
            _repositorio = repositorio;
        }

        [HttpGet]
        public ActionResult Index()
        {
            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }

            var model = new ConsultaLiberacaoCarregamentoViewModel
            {
                TIPOS_VEICULO = _repositorio.ObterTiposVeiculo()
            };

            return View(model);
        }

        [HttpPost]
        public JsonResult Filtrar(ConsultaLiberacaoFiltroRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var patioUsuario = ConverterInt(Session["Patio"]);
            var resultado = _repositorio.Filtrar(request, patioUsuario);

            return Json(new
            {
                success = resultado.SUCESSO,
                message = resultado.MENSAGEM,
                dados = resultado
            });
        }

        [HttpPost]
        public JsonResult CarregarSimultaneos(string protocolo)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var itens = _repositorio.CarregarProtocolosSimultaneos(protocolo);
            return Json(new { success = true, itens });
        }

        [HttpPost]
        public JsonResult Registrar(RegistrarLiberacaoRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            var usuarioId = ConverterInt(Session["UsuarioId"]);
            if (usuarioId <= 0)
            {
                usuarioId = ConverterInt(Session["IdUsuario"]);
            }

            var patioUsuario = ConverterInt(Session["Patio"]);
            var resultado = _repositorio.Registrar(request, usuarioId, patioUsuario);

            return Json(new
            {
                success = resultado.SUCESSO,
                message = resultado.MENSAGEM,
                requerConfirmacaoJanela = resultado.REQUER_CONFIRMACAO_JANELA,
                mensagemConfirmacaoJanela = resultado.MENSAGEM_CONFIRMACAO_JANELA
            });
        }

        private static int ConverterInt(object valor)
        {
            if (valor == null)
            {
                return 0;
            }

            int numero;
            return int.TryParse(valor.ToString(), out numero) ? numero : 0;
        }
    }
}
