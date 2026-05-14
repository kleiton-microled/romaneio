using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class ManutencaoMarcantesArmazemController : DefaultController
    {
        private readonly IManutencaoMarcantesArmazemRepositorio _repositorio;

        public ManutencaoMarcantesArmazemController(IManutencaoMarcantesArmazemRepositorio repositorio)
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

            return View(new ManutencaoMarcantesArmazemViewModel());
        }

        [HttpPost]
        public JsonResult Filtrar(int? lote)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            if (!lote.HasValue || lote.Value <= 0)
            {
                return Json(new { success = false, message = "Informe o lote" });
            }

            var resultado = _repositorio.Filtrar(lote.Value);
            if (!string.IsNullOrEmpty(resultado.MensagemErro))
            {
                return Json(new { success = false, message = resultado.MensagemErro });
            }

            return Json(new
            {
                success = true,
                embalagem = resultado.Embalagem,
                quantidadeReferencia = resultado.QuantidadeReferencia,
                autonumCs = resultado.AutonumCs,
                linhas = resultado.Linhas,
                totalVolumesListados = resultado.TotalVolumesListados,
                totalConfereReferencia = resultado.TotalConfereReferencia
            });
        }

        [HttpPost]
        public JsonResult Salvar(ManutencaoMarcantesArmazemSalvarRequest request)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            if (request == null)
            {
                return Json(new { success = false, message = "Requisicao invalida." });
            }

            if (request.MarcanteAutonum <= 0)
            {
                return Json(new { success = false, message = "Selecione o marcante" });
            }

            if (!request.NovaQuantidade.HasValue)
            {
                return Json(new { success = false, message = "Informe a quantidade" });
            }

            if (request.AutonumCs <= 0)
            {
                return Json(new { success = false, message = "Execute o filtro novamente." });
            }

            var resultado = _repositorio.Salvar(request);
            if (!string.IsNullOrEmpty(resultado.MensagemErro))
            {
                return Json(new { success = false, message = resultado.MensagemErro });
            }

            return Json(new
            {
                success = true,
                message = "Volumes atualizados com sucesso.",
                embalagem = resultado.Embalagem,
                quantidadeReferencia = resultado.QuantidadeReferencia,
                autonumCs = resultado.AutonumCs,
                linhas = resultado.Linhas,
                totalVolumesListados = resultado.TotalVolumesListados,
                totalConfereReferencia = resultado.TotalConfereReferencia
            });
        }
    }
}
