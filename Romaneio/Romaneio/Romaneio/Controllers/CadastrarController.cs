using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Romaneio.Interfaces;
using Romaneio.Models;
using Romaneio.Classes;
using Romaneio.Extensions;
using Romaneio.Models.ViewModels;

namespace Romaneio.Controllers
{
    public class CadastrarController : Controller
    {
        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly IRomaneioRepositorio _RomaneioRepositorio;

        public CadastrarController(IDadosRepositorio dadossrepositorio, IRomaneioRepositorio romaneiorepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _RomaneioRepositorio = romaneiorepositorio;
        }


        [HttpGet]
        public ActionResult Cadastrar(int cntr = 0, int lote = 0, int item = 0, int AUTONUM_ROMANEIO = 0 )
        {
            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {

                LotesDTO Dados = _RomaneioRepositorio.ObterDadosRomaneio(cntr, lote, item, AUTONUM_ROMANEIO);
                if (Dados == null)
                {
                    Dados = _RomaneioRepositorio.ObterDadosLote(cntr, lote, item);
                }

                

                if (item != 0 || AUTONUM_ROMANEIO != 0)
                {
                    LoteViewModel view = new LoteViewModel
                    {
                        AUTONUM_ROMANEIO = Dados.AUTONUM_ROMANEIO,
                        LISTA_LOTES = _RomaneioRepositorio.ListaLotes(Dados.AUTONUM_CNTR).ToList(),
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
                        FLAG_HUB = Dados.FLAG_HUB,
                        LISTA_CNTR_DESTINO = _DadosRepositorio.ConteinersDestino(Dados.PATIO).ToList(),
                        AUTONUM_CNTR_DESTINO = Dados.AUTONUM_CNTR_DESTINO,

                };
                    if (view.AUTONUM_CNTR_DESTINO > 0)
                    {

                        view.FLAG_CD = true;

                    }

                    LotesDTO DadosHUB = _DadosRepositorio.ObterDadosHubLote(Dados.LOTE);
                    if (DadosHUB != null)
                    {
                        view.HUB = DadosHUB.HUB;
                        view.CIDADE_HUB = DadosHUB.CIDADE_HUB;
                        view.FLAG_HUB = true;
                    }

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

                        view.DESOVA = _DadosRepositorio.ConsultarDesovaCNTR(AUTONUM_ROMANEIO.ToString(),cntr.ToString());

                    }
                    return View(view);
                }
                else
                {

                    LoteViewModel view = new LoteViewModel
                    {
                        AUTONUM_ROMANEIO = Dados.AUTONUM_ROMANEIO,
                        LISTA_LOTES = _RomaneioRepositorio.ListaLotes(Dados.AUTONUM_CNTR).ToList(),
                        LOTE = Dados.LOTE,
                        AUTONUM_CNTR = Dados.AUTONUM_CNTR,
                        ID_CONTEINER = Dados.ID_CONTEINER,
                        LISTA_ITENS = _DadosRepositorio.ListaItens(Dados.LOTE, Dados.AUTONUM_CNTR).ToList(),
                        ITEM = Dados.ITEM,
                        TIPOS_SOLICITANTES = _DadosRepositorio.TiposSolicitantes().ToList(),
                        TIPOS_GENERO = _DadosRepositorio.TiposGeneros().ToList(),
                        TIPOS_EMBALAGENS = _DadosRepositorio.TiposEmbalagens().ToList(),
                        TIPOS_DESOVA = _DadosRepositorio.TiposDesova().ToList(),
                        LISTA_CNTR_DESTINO = _DadosRepositorio.ConteinersDestino(Dados.PATIO).ToList(),

                };

                    return View(view);
                }

                

            }
        }



        [HttpGet]
        public ActionResult CadastrarAcrescimo(int cntr)
        {
            string lote = _RomaneioRepositorio.InserirLote(cntr);

            return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote });

        }


        [HttpPost]
        public ActionResult SalvarDados(LoteViewModel DadosEntrada, List<Avarias> items)
        {
            try
            {

                if (DadosEntrada.TERMINO != null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: Lote já finalizado!";
                    return RedirectToAction(nameof(Cadastrar), new { cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM });
                }
                else
                {
                    string result = "";

                    result = _RomaneioRepositorio.SalvarDados(DadosEntrada);

                    

                    if (DadosEntrada.AUTONUM_ROMANEIO == 0)
                    {
                        LotesDTO Consulta = _RomaneioRepositorio.ObterDadosRomaneio(DadosEntrada.AUTONUM_ROMANEIO);
                        if (Consulta != null)
                        {
                            LoteViewModel view = new LoteViewModel
                            {
                                AUTONUM_ROMANEIO = Consulta.AUTONUM_ROMANEIO,
                            };

                            DadosEntrada.AUTONUM_ROMANEIO = Consulta.AUTONUM_ROMANEIO;

                        }

                    }
                    
                    if (DadosEntrada.FL_SEM_AVARIA.ToInt() == 1)
                    {
                        _RomaneioRepositorio.ExcluirTodaAvaria(DadosEntrada.AUTONUM_ROMANEIO);
                    }



                    if (result == "Informações salvas com sucesso!")
                    {
                        TempData["MensagemSucesso"] = "Informações salvas com sucesso!";
                    }
                    else
                    {
                        TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + result;
                    }

                    return RedirectToAction(nameof(Cadastrar), new { cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM });
                }

            }
            catch (Exception erro)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + erro;
                return RedirectToAction(nameof(Cadastrar), new { cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM });
            }
        }

        [HttpGet]
        public ActionResult Finalizar(int patio, int lote, string termino, int autonum_romaneio, int cntr, int item)
        {
            try
            {

                if (termino != null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: Lote já finalizado!";
                    return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote, item = item });
                }
                else
                {
                    if (_RomaneioRepositorio.ContaAvarias(autonum_romaneio) == 0 )
                    {
                        TempData["MensagemErro"] = "Nenhuma avaria lançada!"  ;
                        return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote, item = item });
                    }

                    if (_RomaneioRepositorio.VerificarCubagem(autonum_romaneio) == 1)
                    {
                        TempData["MensagemErro"] = "CLA resulta em cubagem superior ao permitido!";
                       // return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote, item = item });
                    }

                    string result = _RomaneioRepositorio.Finalizar(autonum_romaneio);

                    if (result == "Finalizado com sucesso!")
                    {
                        TempData["MensagemSucesso"] = "Finalizado com sucesso!";
                        return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote, item = item });
                    }
                    else
                    {
                        TempData["MensagemErro"] = "Erro ao realizar salvar informações" + result;
                        return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote, item = item });
                    }
                }

            }
            catch (Exception erro)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + erro;
                return RedirectToAction(nameof(Cadastrar), new { cntr = cntr, lote = lote, item = item });
            }
        }


        public ActionResult Posicionamento(string patio, string cntr, string lote, string item, int romaneio)
        {

            IEnumerable<Posicionamento> tabela = _DadosRepositorio.ConsultarPosicoesItem(patio, cntr, lote, item);

            PosicionamentoViewModel DadosPosicionamento = new PosicionamentoViewModel
            {
                ITEM = item.ToInt(),
                QUANTIDADE = "",
                POSICAO = "",
                ARMAZEM = "",
                PATIO = patio,
                LOTE = lote.ToInt(),
                AUTONUM_CNTR = cntr.ToInt(),
                LISTA_ARMAZENS = _DadosRepositorio.ListaArmazens(patio).ToList(),
                LISTA_POSICIONAMENTOS = tabela.ToList(),
                AUTONUM_CS = _DadosRepositorio.ObterAutonumCS(patio, cntr, lote, item),
                AUTONUM_ROMANEIO = romaneio
            };

            return View("Posicionamento", DadosPosicionamento);

        }


        public ActionResult SalvarYard(PosicionamentoViewModel DadosEntrada)
        {
            try
            {
                if (DadosEntrada.QUANTIDADE == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: Informe a quantidade!";
                    return RedirectToAction(nameof(Posicionamento), new { patio = DadosEntrada.PATIO, cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM, romaneio = DadosEntrada.AUTONUM_ROMANEIO });

                }
                else if (DadosEntrada.POSICAO == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: Informe a posição!";
                    return RedirectToAction(nameof(Posicionamento), new { patio = DadosEntrada.PATIO, cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM, romaneio = DadosEntrada.AUTONUM_ROMANEIO });
                }
                else if (DadosEntrada.ARMAZEM == null)
                {
                    TempData["MensagemErro"] = "Erro ao realizar salvar informações: Informe o armazem!";
                    return RedirectToAction(nameof(Posicionamento), new { patio = DadosEntrada.PATIO, cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM, romaneio = DadosEntrada.AUTONUM_ROMANEIO });
                }
                else
                {
                    var result = _DadosRepositorio.InsereYard(DadosEntrada.AUTONUM_CS, DadosEntrada.QUANTIDADE, DadosEntrada.ARMAZEM, DadosEntrada.POSICAO);
                    TempData["MensagemSucesso"] = "Salvo com sucesso!";
                    return RedirectToAction(nameof(Posicionamento), new { patio = DadosEntrada.PATIO, cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM, romaneio = DadosEntrada.AUTONUM_ROMANEIO });

                }

            }
            catch (Exception erro)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações:" + erro;
                return RedirectToAction(nameof(Posicionamento), new { patio = DadosEntrada.PATIO, cntr = DadosEntrada.AUTONUM_CNTR, lote = DadosEntrada.LOTE, item = DadosEntrada.ITEM, romaneio = DadosEntrada.AUTONUM_ROMANEIO });
            }
        }


        [HttpPost]
        public JsonResult ExcluirYard(int AUTONUM, int AUTONUM_CS)
        {
            if (AUTONUM != null)
            {
                _DadosRepositorio.DeletaYard(AUTONUM, AUTONUM_CS);
                return Json(new { success = true });
            }
            return Json(new { success = false, message = "Item não encontrado" });
        }

      
        

        public ActionResult CarregaLacresRO(int autonum_romaneio, int patio, int cntr, int lote, string item, int autonum_cntr_destino)
        { 

            LoteViewModel dados = new LoteViewModel
            {
                LISTA_LACRES = ObterLacres(autonum_romaneio).ToList(),
                AUTONUM_ROMANEIO = autonum_romaneio,
                PATIO = patio,
                LOTE = lote.ToInt(),
                AUTONUM_CNTR = cntr.ToInt(),
                ITEM = item,
                TERMINO = "",
                AUTONUM_CNTR_DESTINO = autonum_cntr_destino
             };


            return PartialView("Lacres", dados);

        }
        private IEnumerable<Lacres> ObterLacres(int AUTONUM_ROMANEIO)
        {
            return _RomaneioRepositorio.CarregaLacresRO(AUTONUM_ROMANEIO);
        }

        [HttpPost]
        public ActionResult InsereLacreRO(LoteViewModel dados)
        {
            if (dados.AUTONUM_ROMANEIO == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: armazem não encontrado!";
                return RedirectToAction(nameof(CarregaLacresRO), new { autonum_romaneio = dados.AUTONUM_ROMANEIO, patio= dados.PATIO, cntr = dados.AUTONUM_CNTR, lote = dados.LOTE, item = dados.ITEM, autonum_cntr_destino = dados.AUTONUM_CNTR_DESTINO });
            }
            if (dados.AUTONUM_CNTR_DESTINO == null)
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: armazem não encontrado!";
                return RedirectToAction(nameof(CarregaLacresRO), new { autonum_romaneio = dados.AUTONUM_ROMANEIO, patio = dados.PATIO, cntr = dados.AUTONUM_CNTR, lote = dados.LOTE, item = dados.ITEM, autonum_cntr_destino = dados.AUTONUM_CNTR_DESTINO });
            }
            if (dados.LACRE == "")
            {
                TempData["MensagemErro"] = "Erro ao realizar salvar informações: Lacre obrigatorio!";
                return RedirectToAction(nameof(CarregaLacresRO), new { autonum_romaneio = dados.AUTONUM_ROMANEIO, patio = dados.PATIO, cntr = dados.AUTONUM_CNTR, lote = dados.LOTE, item = dados.ITEM, autonum_cntr_destino = dados.AUTONUM_CNTR_DESTINO });
            }

            _RomaneioRepositorio.InsereLacreRO(dados.AUTONUM_ROMANEIO, dados.AUTONUM_CNTR_DESTINO, dados.LACRE);

            LoteViewModel dadosAtualizados = new LoteViewModel
            {
                LISTA_LACRES = ObterLacres(dados.AUTONUM_ROMANEIO).ToList(),
                AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO,
                PATIO = dados.PATIO,
                LOTE = dados.LOTE.ToInt(),
                AUTONUM_CNTR = dados.AUTONUM_CNTR.ToInt(),
                AUTONUM_CNTR_DESTINO = dados.AUTONUM_CNTR_DESTINO
            };


            return PartialView("Lacres", dadosAtualizados);
        }

        [HttpPost]
        public JsonResult ExcluirLacreRO(int AUTONUM_LACRE)
        {
            if (Session["Logado"] == null)
            {
                return Json(new { success = false, message = "Sessao caiu" });
            }
            else
            {
                if (AUTONUM_LACRE != null)
                {
                    _RomaneioRepositorio.ExcluirLacreRO(AUTONUM_LACRE);
                    return Json(new { success = true });
                }
                return Json(new { success = false, message = "Item não encontrado" });
            }
        }


    }
}