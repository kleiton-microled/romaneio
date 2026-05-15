using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class InventarioCargaSoltaArmazemController : DefaultController
    {
        private readonly IInventarioCargaSoltaArmazemRepositorio _repositorio;

        public InventarioCargaSoltaArmazemController(IInventarioCargaSoltaArmazemRepositorio repositorio)
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

            var model = new InventarioCargaSoltaArmazemViewModel
            {
                PATIO = ConverterInt(Session["Patio"]),
                COD_USUARIO = ConverterInt(Session["UsuarioId"])
            };

            return View(model);
        }

        [HttpGet]
        public JsonResult ObterCombos()
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            int patio = ConverterInt(Session["Patio"]);
            return JsonGet(new
            {
                success = true,
                armazens = _repositorio.ListarArmazens(patio),
                motivos = _repositorio.ListarMotivos()
            });
        }

        [HttpGet]
        public JsonResult ObterPorMarcante(int id)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            int patio = ConverterInt(Session["Patio"]);
            var dados = _repositorio.ObterEstadoMarcante(id, patio);
            return JsonGet(new { success = dados.SUCESSO, dados });
        }

        [HttpGet]
        public JsonResult ObterPorLote(long lote, int marcante, long autonumCsYard = 0)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            var dados = _repositorio.ObterEstadoLote(lote, marcante, autonumCsYard);
            return JsonGet(new { success = dados.SUCESSO, dados });
        }

        [HttpGet]
        public JsonResult ObterItem(string idGravacao)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            var r = _repositorio.ObterDetalheItem(idGravacao);
            return JsonGet(new { success = r.SUCESSO, message = r.MENSAGEM, qtde = r.QTDE, embalagem = r.EMBALAGEM, local = r.LOCAL, motivoPadrao = r.MOTIVO_PADRAO });
        }

        [HttpGet]
        public JsonResult Historico(int marcante)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            return JsonGet(new
            {
                success = true,
                linhas = _repositorio.ListarHistorico(marcante)
            });
        }

        [HttpGet]
        public JsonResult ResolverPrateleira(string texto)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            int patio = ConverterInt(Session["Patio"]);
            var r = _repositorio.ResolverPrateleira(texto ?? string.Empty, patio);
            return JsonGet(new
            {
                success = r.SUCESSO,
                message = r.MENSAGEM,
                armazemId = r.ARMAZEM_ID,
                yard = r.YARD,
                yardTag = r.YARD_TAG,
                modoContener = r.MODO_CONTENER
            });
        }

        [HttpPost]
        public JsonResult Salvar()
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            string raw;
            using (var reader = new System.IO.StreamReader(Request.InputStream))
            {
                raw = reader.ReadToEnd();
            }

            InventarioCsSalvarRequest request;
            try
            {
                request = Newtonsoft.Json.JsonConvert.DeserializeObject<InventarioCsSalvarRequest>(raw);
            }
            catch
            {
                return Json(new { success = false, message = "Payload invalido" });
            }

            if (request == null)
            {
                return Json(new { success = false, message = "Payload invalido" });
            }

            int codUsuario = ConverterInt(Session["UsuarioId"]);
            int patio = ConverterInt(Session["Patio"]);
            var resultado = _repositorio.Salvar(request, codUsuario, patio);
            return Json(new
            {
                success = resultado.SUCESSO,
                message = resultado.MENSAGEM,
                marcanteSugerido = resultado.MARCANTE_SUGERIDO_CONFLITO
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

        private JsonResult JsonGet(object data)
        {
            return Json(data, JsonRequestBehavior.AllowGet);
        }
    }
}
