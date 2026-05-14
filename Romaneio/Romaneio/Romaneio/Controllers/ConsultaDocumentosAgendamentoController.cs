using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.IO;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    /// <summary>
    /// Modulo web equivalente ao executavel ConsultaDocColetor / ConsultaDocumentos (WinForms).
    /// Abrir em nova aba ou iframe: /ConsultaDocumentosAgendamento/Index?ltl=0&amp;lote=LOTE&amp;ag=AUTONUM_TB_AG_CS
    /// LTL: /ConsultaDocumentosAgendamento/Index?ltl=1&amp;lote=LOTE (resolve TB_SOLICITACAO_LTL.ID no servidor).
    /// </summary>
    public class ConsultaDocumentosAgendamentoController : DefaultController
    {
        private readonly IConsultaDocumentosAgendamentoRepositorio _repositorio;

        public ConsultaDocumentosAgendamentoController(IConsultaDocumentosAgendamentoRepositorio repositorio)
        {
            _repositorio = repositorio;
        }

        [HttpGet]
        public ActionResult Index(bool ltl = false, int lote = 0, int ag = 0, int? idLtl = null)
        {
            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }

            int usuarioId = ConverterInt(Session["UsuarioId"]);
            if (lote <= 0)
            {
                return View("Index", new ConsultaDocumentosAgendamentoPageViewModel
                {
                    LTL = ltl,
                    LOTE = lote,
                    AUTONUM_AGENDAMENTO = ag,
                    ID_SOLICITACAO_LTL = idLtl,
                    TITULO = "Documentos do agendamento",
                    Erro = "Informe o lote (BL) valido."
                });
            }

            int? idSolicitacao = idLtl;
            if (ltl)
            {
                idSolicitacao = idSolicitacao ?? _repositorio.ResolverIdSolicitacaoLtlPorLote(lote);
                if (!idSolicitacao.HasValue || idSolicitacao.Value <= 0)
                {
                    return View("Index", new ConsultaDocumentosAgendamentoPageViewModel
                    {
                        LTL = true,
                        LOTE = lote,
                        TITULO = "Documentos LTL",
                        Erro = "Solicitacao LTL nao localizada para o lote informado."
                    });
                }

                return View("Index", new ConsultaDocumentosAgendamentoPageViewModel
                {
                    LTL = true,
                    LOTE = lote,
                    ID_SOLICITACAO_LTL = idSolicitacao,
                    AUTONUM_AGENDAMENTO = idSolicitacao.Value,
                    TITULO = "Documentos LTL - Lote " + lote
                });
            }

            if (ag <= 0)
            {
                return View("Index", new ConsultaDocumentosAgendamentoPageViewModel
                {
                    LTL = false,
                    LOTE = lote,
                    TITULO = "Documentos do agendamento",
                    Erro = "Informe o codigo do agendamento (TB_AG_CS.AUTONUM)."
                });
            }

            try
            {
                _repositorio.InserirBloqueioDocumento(usuarioId, ag);
            }
            catch
            {
                // Bloqueio duplicado ou indisponivel: nao impede visualizacao (comportamento tolerante ao legado).
            }

            return View("Index", new ConsultaDocumentosAgendamentoPageViewModel
            {
                LTL = false,
                LOTE = lote,
                AUTONUM_AGENDAMENTO = ag,
                TITULO = "Documentos - Agendamento " + ag + " / Lote " + lote
            });
        }

        [HttpGet]
        public JsonResult Listar(bool ltl, int lote, int ag)
        {
            if (Session["Logado"] == null)
            {
                return JsonGet(new { success = false, message = "Sessao caiu" });
            }

            if (lote <= 0)
            {
                return JsonGet(new { success = false, message = "Lote invalido" });
            }

            if (ltl)
            {
                int? id = ag > 0 ? ag : _repositorio.ResolverIdSolicitacaoLtlPorLote(lote);
                if (!id.HasValue || id.Value <= 0)
                {
                    return JsonGet(new { success = false, message = "Solicitacao LTL nao localizada" });
                }

                var lista = _repositorio.ListarDocumentosLtl(id.Value, lote);
                return JsonGet(new { success = true, itens = lista });
            }

            if (ag <= 0)
            {
                return JsonGet(new { success = false, message = "Agendamento invalido" });
            }

            var itensCs = _repositorio.ListarDocumentosCargaSolta(ag, lote);
            return JsonGet(new { success = true, itens = itensCs });
        }

        [HttpGet]
        public ActionResult Arquivo(int id, bool ltl, int lote, int ag)
        {
            if (Session["Logado"] == null)
            {
                return new HttpStatusCodeResult(401, "Sessao expirada");
            }

            DocumentoAgendamentoArquivoMetaDto meta = _repositorio.ObterMetaArquivo(id, ag, ltl, lote);
            if (meta == null)
            {
                return HttpNotFound("Documento nao encontrado ou sem permissao de contexto.");
            }

            if (!meta.TEM_BLOB)
            {
                return Content(
                    "Arquivo indisponivel neste modulo: o binario nao esta em TB_AG_DOCUMENTOS.DOCUMENTO. " +
                    "No legado WinForms o arquivo e obtido via servico Bandeirantes.AnexoDiretorio (ADMPatio). " +
                    "Integrar esse servico no Romaneio quando necessario.",
                    "text/plain");
            }

            byte[] bytes = _repositorio.ObterBlobDocumento(id);
            if (bytes == null || bytes.Length == 0)
            {
                return HttpNotFound("Conteudo vazio.");
            }

            string ext = (meta.EXTENSAO ?? string.Empty).Trim().ToUpperInvariant();
            string mime = "application/octet-stream";
            if (ext.Contains("PDF"))
            {
                mime = "application/pdf";
            }
            else if (ext.Contains("JPG") || ext.Contains("JPEG"))
            {
                mime = "image/jpeg";
            }
            else if (ext.Contains("PNG"))
            {
                mime = "image/png";
            }
            else if (ext.Contains("GIF"))
            {
                mime = "image/gif";
            }

            return File(bytes, mime);
        }

        [HttpPost]
        public JsonResult Fechar(int ag)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }

            int usuarioId = ConverterInt(Session["UsuarioId"]);
            if (ag > 0 && usuarioId > 0)
            {
                _repositorio.RemoverBloqueioDocumento(usuarioId, ag);
            }

            return Json(new { success = true });
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
