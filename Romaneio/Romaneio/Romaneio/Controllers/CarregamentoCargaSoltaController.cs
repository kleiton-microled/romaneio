using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class CarregamentoCargaSoltaController : DefaultController
    {
        private readonly ICarregamentoCargaSoltaRepositorio _repositorio;

        public CarregamentoCargaSoltaController(ICarregamentoCargaSoltaRepositorio repositorio)
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

            CarregamentoCargaSoltaViewModel model = new CarregamentoCargaSoltaViewModel
            {
                PATIO = patioUsuario,
                BIPAGEM_OBRIGATORIA = ConverterInt(Session["FlagCarregamentoMarcanteBip"]) == 1,
                BIPAGEM_BLOQUEAR_APOS_VALIDAR = ConverterInt(Session["FlagCarregamentoMarcanteBip"]) == 1,
                VEICULOS = _repositorio.ListarVeiculos(patioUsuario)
            };

            return View(model);
        }

        [HttpGet]
        public JsonResult ListarVeiculos()
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            int patioUsuario = ConverterInt(Session["Patio"]);
            return JsonGet(new
            {
                success = true,
                veiculos = _repositorio.ListarVeiculos(patioUsuario)
            });
        }

        [HttpGet]
        public JsonResult ListarOrdens(string placa)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            return JsonGet(new
            {
                success = true,
                ordens = _repositorio.ListarOrdens(placa)
            });
        }

        [HttpGet]
        public JsonResult ListarCarregamentos(string placa, int autonumCs)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            return JsonGet(new
            {
                success = true,
                marcantes = _repositorio.ListarCarregamentos(placa, autonumCs)
            });
        }

        [HttpGet]
        public JsonResult ListarLocais(int autonumCs)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            return JsonGet(new
            {
                success = true,
                locais = _repositorio.ListarLocais(autonumCs)
            });
        }

        [HttpGet]
        public JsonResult BuscarMarcante(int marcante)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            MarcanteDadosDto dados = _repositorio.BuscarMarcante(marcante);
            if (dados == null)
            {
                return JsonGet(new { success = false, message = "Marcante nao informado/invalido" });
            }

            return JsonGet(new { success = true, dados });
        }

        [HttpPost]
        public JsonResult Carregar(CarregarItemRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            int codUsuario = ConverterInt(Session["UsuarioId"]);
            CarregarItemResult resultado = _repositorio.Carregar(request, codUsuario);

            return Json(new
            {
                success = resultado.SUCESSO,
                parcial = resultado.PARCIAL,
                message = resultado.MENSAGEM
            });
        }

        [HttpPost]
        public JsonResult Cancelar(CancelarItemRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            int codUsuario = ConverterInt(Session["UsuarioId"]);
            CancelarItemResult resultado = _repositorio.Cancelar(request, codUsuario);

            return Json(new
            {
                success = resultado.SUCESSO,
                message = resultado.MENSAGEM
            });
        }

        [HttpGet]
        public JsonResult Documento(int loteBl)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            DocumentoOcResult resultado = _repositorio.ObterDocumento(loteBl);
            return JsonGet(new
            {
                success = resultado.SUCESSO,
                ltl = resultado.LTL,
                autonumAgCs = resultado.AUTONUM_AG_CS,
                loteBl = resultado.LOTE_BL,
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

        private JsonResult JsonGet(object data)
        {
            return Json(data, JsonRequestBehavior.AllowGet);
        }
    }
}
