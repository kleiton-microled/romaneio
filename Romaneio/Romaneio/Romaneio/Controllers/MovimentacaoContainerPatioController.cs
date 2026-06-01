using System;
using System.Linq;
using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;

namespace Romaneio.Controllers
{
    public class MovimentacaoContainerPatioController : DefaultController
    {
        private readonly IMovimentacaoContainerPatioRepositorio _repositorio;

        public MovimentacaoContainerPatioController(IMovimentacaoContainerPatioRepositorio repositorio)
        {
            _repositorio = repositorio;
        }

        [HttpGet]
        public ActionResult Index()
        {
            if (Session["Logado"] == null)
                return RedirectToAction("Index", "Home");

            var patioUsuario = ConverterInt(Session["Patio"]);
            var usuarioId = ConverterInt(Session["UsuarioId"]);
            var perm = _repositorio.ObterPermissoesAvarias(usuarioId);

            var model = new MovimentacaoContainerPatioViewModel
            {
                PATIO = patioUsuario,
                DESCR_PATIO = _repositorio.ObterDescricaoPatio(patioUsuario),
                MOTIVOS = _repositorio.ListarMotivosAtivos(),
                CAMERAS = _repositorio.ListarCameras(patioUsuario),
                PERMISSOES_AVARIAS = perm
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
        public JsonResult CarregarDetalhe(string idConteiner)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var usuarioId = ConverterInt(Session["UsuarioId"]);
            var dto = _repositorio.CarregarDetalhe(idConteiner, patioUsuario, usuarioId);

            if (!string.IsNullOrEmpty(dto.TEXTO_CONTEXTUAL) && dto.TEXTO_CONTEXTUAL.StartsWith("__ERRO__|", StringComparison.Ordinal))
            {
                var msg = dto.TEXTO_CONTEXTUAL.Substring("__ERRO__|".Length);
                return Json(new { success = false, message = msg });
            }

            return Json(new { success = true, data = dto });
        }

        [HttpPost]
        public JsonResult ProximosMovimentos(string idConteiner)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var list = _repositorio.ProximosMovimentos(idConteiner);
            return Json(new { success = true, dados = list });
        }

        [HttpPost]
        public JsonResult VizinhancaPilha(VizinhancaPilhaRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var list = _repositorio.VizinhancaPilha(request, patioUsuario);
            return Json(new { success = true, dados = list });
        }

        [HttpPost]
        public JsonResult CameraPadrao(string yardDestino)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var id = _repositorio.ObterCameraPadraoPorYard(patioUsuario, yardDestino);
            return Json(new { success = true, cameraId = id });
        }

        [HttpPost]
        public JsonResult ConfirmarMovimentacao(MovimentacaoCntrConfirmRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            var patioUsuario = ConverterInt(Session["Patio"]);
            var usuarioId = ConverterInt(Session["UsuarioId"]);
            var r = _repositorio.ConfirmarMovimentacao(request, patioUsuario, usuarioId);

            return Json(new
            {
                success = r.SUCESSO,
                message = r.MENSAGEM,
                codigo = r.CODIGO,
                exigeCamera = r.EXIGE_CAMERA,
                exigeConfirmacaoSaida = r.EXIGE_CONFIRMACAO_SAIDA,
                exigeVeiculoServ = r.EXIGE_VEICULO_SERV
            });
        }

        [HttpGet]
        public JsonResult HistoricoShifting(string idConteiner, string sistema)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" }, JsonRequestBehavior.AllowGet);

            var list = _repositorio.HistoricoShifting(idConteiner, sistema);
            return Json(new { success = true, dados = list }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult AvariasGrid(long autonumCntr)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" }, JsonRequestBehavior.AllowGet);

            var list = _repositorio.ListarAvariasAbertas(autonumCntr);
            return Json(new
            {
                success = true,
                dados = list,
                finalizada = _repositorio.AvariasFinalizadas(autonumCntr)
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult AvariasHistorico(long autonumCntr)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" }, JsonRequestBehavior.AllowGet);

            var list = _repositorio.HistoricoAvarias(autonumCntr);
            return Json(new { success = true, dados = list }, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult OpcoesAvarias()
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" }, JsonRequestBehavior.AllowGet);

            return Json(new
            {
                success = true,
                locais = _repositorio.ListarDescricoesAvariasLocal(),
                tipos = _repositorio.ListarDescricoesAvariasTipo()
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult AvariasIncluir(AvariaIncluirRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            try
            {
                var usuarioId = ConverterInt(Session["UsuarioId"]);
                if (_repositorio.AvariasFinalizadas(request.AUTONUM_CNTR))
                    return Json(new { success = false, message = "Ja consta finalizacao do Patio para a unidade" });

                if (request.TIPOS_DESCR == null || !request.TIPOS_DESCR.Any())
                    return Json(new { success = false, message = "Nenhum tipo selecionado" });
                if (string.IsNullOrWhiteSpace(request.LOCAL_DESCR))
                    return Json(new { success = false, message = "Nenhum Local selecionado" });

                foreach (var tipo in request.TIPOS_DESCR.Where(t => !string.IsNullOrWhiteSpace(t)))
                    _repositorio.IncluirAvaria(request.AUTONUM_CNTR, request.LOCAL_DESCR.Trim(), tipo.Trim(), usuarioId);
                return Json(new { success = true, message = "Avaria registrada" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult AvariasExcluir(AvariaExcluirRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            if (request.AUTONUM_AVARIA <= 0)
                return Json(new { success = false, message = "Avaria apontada pelo Operador" });

            if (_repositorio.AvariasFinalizadas(request.AUTONUM_CNTR))
                return Json(new { success = false, message = "Ja consta finalizacao do Patio para a unidade" });

            try
            {
                _repositorio.ExcluirAvaria(request.AUTONUM_AVARIA);
                return Json(new { success = true, message = "Avaria excluida" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public JsonResult AvariasFinalizar(AvariaFinalizarRequest request)
        {
            if (Session["Logado"] == null)
                return Json(new { success = false, message = "Sessao caiu" });

            if (_repositorio.AvariasFinalizadas(request.AUTONUM_CNTR))
                return Json(new { success = false, message = "Ja consta finalizacao do Patio para a unidade" });

            try
            {
                _repositorio.FinalizarAvarias(request.AUTONUM_CNTR);
                return Json(new { success = true, message = "Finalizacao registrada" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static int ConverterInt(object valor)
        {
            if (valor == null) return 0;
            int n;
            return int.TryParse(valor.ToString(), out n) ? n : 0;
        }
    }
}
