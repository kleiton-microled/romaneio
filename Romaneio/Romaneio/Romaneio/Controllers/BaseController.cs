using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Romaneio.Models;

namespace Romaneio.Controllers
{
    public class BaseController : Controller
    {
        // GET: Base
        public bool Validar<T>(Entidade<T> entidade)
        {
            ModelState.Clear();

            entidade.Validar();

            foreach (var erro in entidade.ValidationResult.Errors)
                ModelState.AddModelError(erro.PropertyName, erro.ErrorMessage);

            return entidade.Valido;
        }

        public ActionResult RetornarErros()
        {
            Response.StatusCode = (int)HttpStatusCode.BadRequest;

            return Json(new
            {
                erros = ModelState.Values.SelectMany(v => v.Errors)
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult RetornarErro(string mensagem)
        {
            Response.StatusCode = (int)HttpStatusCode.BadRequest;

            return Json(mensagem, JsonRequestBehavior.AllowGet);
        }

        public class ResponseJson
        {
            public ResponseJson() { }
            public string statusRetorno { get; set; }

            public bool possuiDados { get; set; }

            public object objetoRetorno { get; set; }

            public string Mensagem { get; set; }
        }

        public ResponseJson retornoJson = new ResponseJson();
    }
}