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
    public class FechamentosController : Controller
    {
        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly IFechamentosRepositorio _FechamentosRepositorio;


        public FechamentosController(IDadosRepositorio dadossrepositorio, IFechamentosRepositorio fechamentosrepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _FechamentosRepositorio = fechamentosrepositorio;

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
                    totalItems = _FechamentosRepositorio.ConsultarFechamentos(PATIO).Where(p => p.ID_CONTEINER.Contains(filtro.ID_CONTEINER_FILTRO)).Count();

                }
                else
                {
                    totalItems = _FechamentosRepositorio.ConsultarFechamentos(PATIO).Count();
                }



                // Calculando o total de páginas
                int totalPages = (int)Math.Ceiling((double)totalItems / pageSize);

                // Pegando os itens da página atual
                List<Conteiners> items;

                if (!string.IsNullOrEmpty(filtro.ID_CONTEINER_FILTRO))
                {

                    items = _FechamentosRepositorio.ConsultarFechamentos(PATIO).Where(p => p.ID_CONTEINER.Contains(filtro.ID_CONTEINER_FILTRO))
                                              .Skip((page - 1) * pageSize)
                                              .Take(pageSize)
                                              .ToList();

                }
                else
                {
                    items = _FechamentosRepositorio.ConsultarFechamentos(PATIO)
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

                return View("Consultar", model);


            }
        }

        public ActionResult Lotes(string cntr)
        {
            cntr = cntr.Replace("div-", "");

            IEnumerable<LotesDTO> dados = _FechamentosRepositorio.ConsultarFechamentosLotesConteiners(cntr);

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



        [HttpGet]
        public ActionResult ConsultarCadastro(int cntr = 0, int lote = 0, int item = 0, int AUTONUM_ROMANEIO = 0)
        {
            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                LotesDTO Dados = _FechamentosRepositorio.ObterDadosRomaneioFechado(cntr, lote, item, AUTONUM_ROMANEIO);
                if (Dados == null)
                {
                    Dados = _FechamentosRepositorio.ObterDadosLote(cntr, lote, item);
                }

                if (item != 0 || AUTONUM_ROMANEIO != 0)
                {
                    LoteViewModel view = new LoteViewModel
                    {
                        AUTONUM_ROMANEIO = Dados.AUTONUM_ROMANEIO,
                        LISTA_LOTES = _FechamentosRepositorio.ListaLotesFechados(Dados.AUTONUM_CNTR).ToList(),
                        LOTE = Dados.LOTE,
                        AUTONUM_CNTR = Dados.AUTONUM_CNTR,
                        ID_CONTEINER = Dados.ID_CONTEINER,
                        INICIO = Dados.INICIO,
                        DEMURRAGE = Dados.DEMURRAGE,
                        TERMINO = Dados.TERMINO,
                        CONFERENTE = Dados.CONFERENTE,
                        FISCAL = Dados.FISCAL,
                        SOLICITANTE = Dados.SOLICITANTE,
                        TIPOS_SOLICITANTES = _DadosRepositorio.TiposSolicitantes().ToList(),
                        DESOVA = Dados.DESOVA,
                        TIPOS_DESOVA = _DadosRepositorio.TiposDesova().ToList(),
                        INSPECAO_MAPA = Dados.INSPECAO_MAPA,
                        VOLUME = Dados.VOLUME,
                        PESO_APURADO = Dados.PESO_APURADO,
                        PESO_BRUTO = Dados.PESO_BRUTO,
                        IMO = Dados.IMO,
                        UNDG = Dados.UNDG,
                        GENERO = Dados.GENERO,
                        TIPOS_GENERO = _DadosRepositorio.TiposGeneros().ToList(),
                        LISTA_ITENS = _DadosRepositorio.ListaItens(Dados.LOTE, Dados.AUTONUM_CNTR).ToList(),
                        ITEM = Dados.ITEM,
                        QUANTIDADE = Dados.QUANTIDADE,
                        EMBALAGEM = Dados.EMBALAGEM,
                        TIPOS_EMBALAGENS = _DadosRepositorio.TiposEmbalagens().ToList(),
                        MERCADORIA = Dados.MERCADORIA,
                        MARCA = Dados.MARCA,
                        OBS = Dados.OBS,
                        FL_OUTRA_AVARIA = Dados.FL_OUTRA_AVARIA,
                        DESCR_OUTRA_AVARIA = Dados.DESCR_OUTRA_AVARIA,
                        FL_FALTA = Dados.FL_FALTA,
                        DESCR_FALTA = Dados.DESCR_FALTA,
                        FL_ACRESCIMO = Dados.FL_ACRESCIMO,
                        DESCR_ACRESCIMO = Dados.DESCR_ACRESCIMO,
                        FL_AGUARDANDO_RECONHECIMENTO = Dados.FL_AGUARDANDO_RECONHECIMENTO,
                        DESCR_AGUARDANDO_RECONHECIMENTO = Dados.DESCR_AGUARDANDO_RECONHECIMENTO,
                        LOCALIZACAO = Dados.LOCALIZACAO,
                        LOCAL = Dados.LOCAL,
                        SUB_CLASSE = Dados.SUB_CLASSE,
                        DESCRICAO = Dados.DESCRICAO,
                        NCM = Dados.NCM,
                        PATIO = Dados.PATIO,
                        EMPRESA = Dados.EMPRESA,
                        FL_SEM_AVARIA = Dados.FL_SEM_AVARIA,

                    };


                    if (view.INICIO == null)
                    {

                        view.INICIO = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

                    }

                    if (view.CONFERENTE == null)
                    {

                        view.CONFERENTE = Session["NomeUsuario"].ToString();

                    }

                    if (view.SOLICITANTE == null)
                    {

                        view.SOLICITANTE = _DadosRepositorio.ConsultarSolicitanteCNTR(AUTONUM_ROMANEIO.ToString(), cntr.ToString());

                    }

                    if (view.DESOVA == null)
                    {

                        view.DESOVA = _DadosRepositorio.ConsultarDesovaCNTR(AUTONUM_ROMANEIO.ToString(), cntr.ToString());

                    }
                    return View(view);
                }
                else
                {
                    Dados.AUTONUM_ROMANEIO = 0;
                    Dados.ITEM = "0";
                    LoteViewModel view = new LoteViewModel
                    {
                        AUTONUM_ROMANEIO = Dados.AUTONUM_ROMANEIO,
                        LISTA_LOTES = _FechamentosRepositorio.ListaLotesFechados(Dados.AUTONUM_CNTR).ToList(),
                        LOTE = Dados.LOTE,
                        AUTONUM_CNTR = Dados.AUTONUM_CNTR,
                        ID_CONTEINER = Dados.ID_CONTEINER,
                        LISTA_ITENS = _DadosRepositorio.ListaItens(Dados.LOTE, Dados.AUTONUM_CNTR).ToList(),
                        ITEM = Dados.ITEM,
                        TIPOS_SOLICITANTES = _DadosRepositorio.TiposSolicitantes().ToList(),
                        TIPOS_GENERO = _DadosRepositorio.TiposGeneros().ToList(),
                        TIPOS_EMBALAGENS = _DadosRepositorio.TiposEmbalagens().ToList(),
                        TIPOS_DESOVA = _DadosRepositorio.TiposDesova().ToList(),
                    };

                    return View(view);
                }

            }
        }

    }
}
