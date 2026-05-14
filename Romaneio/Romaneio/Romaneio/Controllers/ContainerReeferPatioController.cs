using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;

namespace Romaneio.Controllers
{
    public class ContainerReeferPatioController : DefaultController
    {
        private readonly IContainerReeferPatioRepositorio _repositorio;

        public ContainerReeferPatioController(IContainerReeferPatioRepositorio repositorio)
        {
            _repositorio = repositorio;
        }

        [HttpGet]
        public ActionResult Index()
        {
            if (Session["Logado"] == null)
                return RedirectToAction("Index", "Home");

            var patioUsuario = ConverterInt(Session["Patio"]);
            var model = new ContainerReeferPatioViewModel
            {
                PATIO = patioUsuario,
                DESCR_PATIO = _repositorio.ObterDescricaoPatio(patioUsuario)
            };
            return View(model);
        }

        [HttpPost]
        public JsonResult BuscarPorFinal(string final)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var r = _repositorio.BuscarPorFinal(final, patioUsuario);
            return Json(new { success = r.SUCESSO, message = r.MENSAGEM, idConteiner = r.ID_CONTEINER });
        }

        [HttpPost]
        public JsonResult CarregarConteiner(string idConteiner)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var dto = _repositorio.CarregarConteiner(idConteiner, patioUsuario);
            if (!string.IsNullOrEmpty(dto.ERRO))
                return Json(new { success = false, message = dto.ERRO });

            return Json(new { success = true, data = dto });
        }

        [HttpPost]
        public JsonResult ListarHistorico(long autonumIpa, long autonumRdx, long autonumOp)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var list = _repositorio.ListarHistorico(autonumIpa, autonumRdx, autonumOp);
            return Json(new { success = true, dados = list });
        }

        [HttpPost]
        public JsonResult SalvarMonitoramento(SalvarMonitoramentoRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var usuarioId = ConverterInt(Session["UsuarioId"]);
            var r = _repositorio.SalvarMonitoramento(request, patioUsuario, usuarioId);

            return Json(new
            {
                success = r.SUCESSO,
                message = r.MENSAGEM,
                exigeConfirmacaoDivergenciaTemperatura = r.EXIGE_CONFIRMACAO_DIVERGENCIA_TEMPERATURA
            });
        }

        [HttpPost]
        public JsonResult RegistrarPlugOff(long autonumIpa)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var r = _repositorio.RegistrarPlugOff(autonumIpa, patioUsuario);
            return Json(new { success = r.SUCESSO, message = r.MENSAGEM });
        }

        [HttpPost]
        public JsonResult ListarEntradasPrevistas()
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var list = _repositorio.ListarEntradasPrevistas(patioUsuario);
            return Json(new { success = true, dados = list, total = list.Count });
        }

        [HttpPost]
        public JsonResult ListarEstoqueReefer(bool apenasDesligados, bool comAgendamentoSaida, bool posicionados)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var list = _repositorio.ListarEstoqueReefer(patioUsuario, apenasDesligados, comAgendamentoSaida, posicionados);
            return Json(new { success = true, dados = list, total = list.Count });
        }

        [HttpPost]
        public JsonResult ListarUnidadesDesligadas()
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var list = _repositorio.ListarUnidadesDesligadas(patioUsuario);
            return Json(new { success = true, dados = list, total = list.Count });
        }

        private static int ConverterInt(object valor)
        {
            if (valor == null) return 0;
            int n;
            return int.TryParse(valor.ToString(), out n) ? n : 0;
        }
    }
}
