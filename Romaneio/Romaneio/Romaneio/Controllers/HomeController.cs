using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models;
using Romaneio.Classes;
using Romaneio.Extensions;
 

namespace Romaneio.Controllers
{
    public class HomeController : Controller
    {
        private readonly IDadosRepositorio _DadosRepositorio;

        public HomeController(IDadosRepositorio dadossrepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
        }
        public ActionResult Home()
        {
            return View("Home");
        }

        [HttpGet]
        public ActionResult Index()
        {
            Session["Logado"] = null;
 
            LoginViewModel view = new LoginViewModel
            {
                TIPOS_PATIOS = _DadosRepositorio.ConsultarPatios().ToList(),
                PATIO = 0,
            };
            return View(view);
        }

        [HttpPost]
        public ActionResult Entrar(LoginViewModel DadosEntrada)
        {
            try
            {
                Login usuario = _DadosRepositorio.ConsultarUsuario(DadosEntrada.Login);

                DadosEntrada.nome = usuario.NOME;

                if (usuario != null)
                {
                    if (DadosEntrada.PATIO == 0)
                    {
                        DadosEntrada.TIPOS_PATIOS = _DadosRepositorio.ConsultarPatios().ToList();
                        TempData["MensagemErro"] = "Selecione um pátio!";
                        Session["Logado"] = null;
                        return View("index", DadosEntrada);
                    }

                    if (DadosEntrada.SenhaValida(usuario.SENHA))
                    {

                        Session["Logado"] = 1;
                        Session["NomeUsuario"] = usuario.NOME;
                        Session["Patio"] = DadosEntrada.PATIO;

                        //return RedirectToAction("Consultar", "Consultar", new { PATIO = DadosEntrada.PATIO });
                        return RedirectToAction("Home", "Home", new { PATIO = DadosEntrada.PATIO });
                    }
                    else
                    {
                        DadosEntrada.TIPOS_PATIOS = _DadosRepositorio.ConsultarPatios().ToList();
                        TempData["MensagemErro"] = "Senha inválida!";
                        Session["Logado"] = null;
                        return View("index", DadosEntrada);
                    }
                }
                else
                {
                    DadosEntrada.TIPOS_PATIOS = _DadosRepositorio.ConsultarPatios().ToList();
                    TempData["MensagemErro"] = "Usuário não encontrado!";
                    Session["Logado"] = null;
                    return View("index", DadosEntrada);
                }

            }
            catch (Exception erro)
            {
                DadosEntrada.TIPOS_PATIOS = _DadosRepositorio.ConsultarPatios().ToList();
                TempData["MensagemErro"] = "Erro ao realizar login:" + erro;
                Session["Logado"] = null;
                return View("index", DadosEntrada);
            }
        }


    }



}