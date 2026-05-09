using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class LocalizacaoCarregamentosController : DefaultController
    {
        private readonly ILocalizacaoCarregamentosRepositorio _repositorio;

        public LocalizacaoCarregamentosController(ILocalizacaoCarregamentosRepositorio repositorio)
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

            int patioUsuario = ConverterInt(Session["Patio"]);
            var model = new LocalizacaoCarregamentosViewModel
            {
                PATIO = patioUsuario,
                DESCR_PATIO = _repositorio.ObterDescricaoPatio(patioUsuario)
            };

            return View(model);
        }

        [HttpPost]
        public JsonResult Filtrar(LocalizacaoCarregamentosFiltroRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            int patioUsuario = ConverterInt(Session["Patio"]);
            var resultado = _repositorio.Filtrar(request, patioUsuario);

            return Json(new
            {
                success = resultado.SUCESSO,
                message = resultado.MENSAGEM,
                dados = resultado
            });
        }

        [HttpPost]
        public JsonResult MoverSaida(MoverSaidaRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            int patioUsuario = ConverterInt(Session["Patio"]);
            int usuarioId = ConverterInt(Session["UsuarioId"]);
            var resultado = _repositorio.MoverParaSaida(request, patioUsuario, usuarioId);

            return Json(new
            {
                success = resultado.SUCESSO,
                message = resultado.MENSAGEM
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
