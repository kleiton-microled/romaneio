using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models;
using Romaneio.Models.ViewModels;
using Romaneio.Classes;
using Romaneio.Extensions;

namespace Romaneio.Controllers
{
    public class ConsultarController : Controller
    {

        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly IRomaneioRepositorio _RomaneioRepositorio;

        public ConsultarController(IDadosRepositorio dadossrepositorio, IRomaneioRepositorio romaneiorepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _RomaneioRepositorio = romaneiorepositorio;

        }

        [HttpGet]
        public ActionResult Consultar(int PATIO = 0, int page = 1, int pageSize = 10, FiltroViewModel filtro = null)
        {
            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {

                if (PATIO == 0)
                {
                    PATIO = Session["Patio"].ToInt();
                }

                // Total de registros
                var totalItems = 0;

                if (!string.IsNullOrEmpty(filtro.ID_CONTEINER_FILTRO))
                {
                    totalItems = _RomaneioRepositorio.ConsultarConteiners(PATIO).Where(p => p.ID_CONTEINER.Contains(filtro.ID_CONTEINER_FILTRO)).Count();

                }
                else
                {
                    totalItems = _RomaneioRepositorio.ConsultarConteiners(PATIO).Count();
                }



                // Calculando o total de páginas
                int totalPages = (int)Math.Ceiling((double)totalItems / pageSize);

                // Pegando os itens da página atual
                List<Conteiners> items;

                if (!string.IsNullOrEmpty(filtro.ID_CONTEINER_FILTRO))
                {

                    items = _RomaneioRepositorio.ConsultarConteiners(PATIO).Where(p => p.ID_CONTEINER.Contains(filtro.ID_CONTEINER_FILTRO))
                                              .Skip((page - 1) * pageSize)
                                              .Take(pageSize)
                                              .ToList();

                }
                else
                {
                    items = _RomaneioRepositorio.ConsultarConteiners(PATIO)
                                                .Skip((page - 1) * pageSize)
                                                .Take(pageSize)
                                                .ToList();
                }


                // Criando o ViewModel
                var model = new ConsultaViewModel
                {
                    Itens = items,
                    PaginaAtual = page,
                    TotalPaginas = totalPages,
                    TamanhoPagina = pageSize,
                    Patio = PATIO
                };

                return View(model);



            }
        }

        public ActionResult Lotes(string cntr)
        {
            cntr = cntr.Replace("div-", "");

            IEnumerable<LotesDTO> dados = _RomaneioRepositorio.ConsultarLotesConteiners(cntr);

            return PartialView("Lotes", dados); // Retorna um PartialView com os detalhes

        }


        public ActionResult RelacaoUnidades(string lote)
        {

            IEnumerable<RelacaoUnidadesViewModel> dados = _DadosRepositorio.ConsultarUnidades(lote);

            return PartialView("RelacaoUnidades", dados); // Retorna um PartialView com os detalhes

        }

        public ActionResult NCM(string lote)
        {

            IEnumerable<NCMViewModel> dados = _DadosRepositorio.ConsultarNCM(lote);

            return PartialView("NCM", dados); // Retorna um PartialView com os detalhes

        }

        public ActionResult Descricao(string lote)
        {

            DescricaoDTO descr = _DadosRepositorio.ConsultarDescr(lote);


            DescricaoViewModel view = new DescricaoViewModel
            {
                Descricao = descr.Descricao

            };


            return PartialView("Descricao", view);

        }


        public ActionResult Localizacao(string lote)
        {

            LocalizacaoDTO DadosLocalizacao = _DadosRepositorio.ConsultarLocalizacao(lote);

            return PartialView("Localizacao", DadosLocalizacao);

        }
    }
}