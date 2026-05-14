using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;

namespace Romaneio.Controllers
{
    public class DescargaDdcPatioController : DefaultController
    {
        private readonly IDescargaDdcPatioRepositorio _repositorio;

        public DescargaDdcPatioController(IDescargaDdcPatioRepositorio repositorio)
        {
            _repositorio = repositorio;
        }

        [HttpGet]
        public ActionResult Index()
        {
            if (Session["Logado"] == null)
                return RedirectToAction("Index", "Home");

            var patioUsuario = ConverterInt(Session["Patio"]);
            var model = new DescargaDdcPatioViewModel
            {
                PATIO = patioUsuario,
                DESCR_PATIO = _repositorio.ObterDescricaoPatio(patioUsuario)
            };
            return View(model);
        }

        [HttpPost]
        public JsonResult ListarPlacasDdc()
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var list = _repositorio.ListarPlacasDdc(patioUsuario);
            return Json(new { success = true, dados = list });
        }

        [HttpPost]
        public JsonResult ListarEmbalagens()
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var list = _repositorio.ListarEmbalagens();
            return Json(new { success = true, dados = list });
        }

        [HttpPost]
        public JsonResult ResolverGateELotes(long autonumRegistroSaida)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var r = _repositorio.ResolverGateELotes(autonumRegistroSaida, patioUsuario);
            if (r.GATE <= 0)
                return Json(new { success = false, message = "Nao consta entrada para o veiculo informado" });

            return Json(new { success = true, gate = r.GATE, lotes = r.LOTES });
        }

        [HttpPost]
        public JsonResult CarregarCabecalho(long autonumRegistroSaida, long autonumLote)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            if (autonumRegistroSaida <= 0)
                return Json(new { success = false, message = "Placa nao informada" });
            if (autonumLote <= 0)
                return Json(new { success = false, message = "Lote nao informado" });

            var cab = _repositorio.CarregarCabecalho(autonumRegistroSaida, autonumLote, patioUsuario);
            if (cab == null)
                return Json(new { success = false, message = "Nao foi possivel carregar os dados da placa/lote." });

            var itens = _repositorio.ListarItens(cab.CNTR, cab.GATE, patioUsuario);
            return Json(new { success = true, cabecalho = cab, itens = itens });
        }

        [HttpPost]
        public JsonResult ListarItens(long cntr, long gate)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var itens = _repositorio.ListarItens(cntr, gate, patioUsuario);
            return Json(new { success = true, itens = itens });
        }

        [HttpPost]
        public JsonResult SalvarItem(SalvarDescargaDdcRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            request = request ?? new SalvarDescargaDdcRequest();
            var patioUsuario = ConverterInt(Session["Patio"]);
            var usuarioId = ConverterInt(Session["UsuarioId"]);

            var r = _repositorio.SalvarItem(request, patioUsuario, usuarioId);
            return Json(new
            {
                success = r.SUCESSO,
                message = r.MENSAGEM,
                aviso = r.AVISO
            });
        }

        [HttpPost]
        public JsonResult ExcluirItem(long autonumItem)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var r = _repositorio.ExcluirItem(autonumItem, patioUsuario);
            return Json(new { success = r.SUCESSO, message = r.MENSAGEM });
        }

        private static int ConverterInt(object valor)
        {
            if (valor == null) return 0;
            int n;
            return int.TryParse(valor.ToString(), out n) ? n : 0;
        }
    }
}
