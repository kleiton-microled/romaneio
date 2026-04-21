using Romaneio.Extensions;
using Romaneio.Interfaces;
using Romaneio.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace Romaneio.Controllers
{
    public class FotosController : Controller
    {
        private readonly IDadosRepositorio _DadosRepositorio;
        private readonly IRomaneioRepositorio _RomaneioRepositorio;
        public FotosController(IDadosRepositorio dadossrepositorio, IRomaneioRepositorio romaneiorepositorio)
        {
            _DadosRepositorio = dadossrepositorio;
            _RomaneioRepositorio = romaneiorepositorio;
        }


        [HttpGet]
        public ActionResult ConsultarFotos(int AUTONUM_ROMANEIO = 0)
        {

            if (Session["Logado"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                List<FotosRomaneio> dados = _RomaneioRepositorio.ConsultarFotosRomaneio(AUTONUM_ROMANEIO).ToList();

                var model = new FotosRomaneioViewModel
                {
                    AUTONUM_ROMANEIO = AUTONUM_ROMANEIO,
                    FOTOS = dados
                };


                return View("ConsultarFotos", model);
            }
        }

        [HttpGet]
        public ActionResult TirarFoto(int AUTONUM_ROMANEIO)
        {
            FotosRomaneio view = new FotosRomaneio
            {
                AUTONUM_ROMANEIO = AUTONUM_ROMANEIO,
                LISTA_TIPO_FOTO = _RomaneioRepositorio.ListaTipoFoto().ToList(),
            };

            return View(view);
        }

        [HttpPost]
        public ActionResult SalvarFoto(FotosRomaneio dados)
        {

            string diretorio_arquivos = Path.Combine(Server.MapPath(ConfigurationManager.AppSettings["DiretorioFotos"]), dados.AUTONUM_ROMANEIO.ToString());

            // Nome do arquivo com extensão
            var nome_imagem = Guid.NewGuid().ToString() + ".png";

            // Garante que o diretório existe
            Directory.CreateDirectory(diretorio_arquivos);

            // Converte de base64 para bytes
            var imagemBytes = Convert.FromBase64String(dados.FOTO);

            // Caminho completo do arquivo
            var caminhoArquivo = Path.Combine(diretorio_arquivos, nome_imagem);

            // Salva a imagem
            System.IO.File.WriteAllBytes(caminhoArquivo, imagemBytes);

            string caminhoRelativo = Url.Content(Path.Combine(ConfigurationManager.AppSettings["DiretorioFotos"], dados.AUTONUM_ROMANEIO.ToString(), nome_imagem));

            // 8. Exemplo de salvar no banco (ajuste conforme seu modelo/repositorio)
            dados.FOTO = caminhoRelativo;
            dados.ID_TIPO_PROCESSO = 1;

            _RomaneioRepositorio.InsereNovaFoto(dados);

             return RedirectToAction(nameof(ConsultarFotos), new { AUTONUM_ROMANEIO = dados.AUTONUM_ROMANEIO });
        }

        public ActionResult VisualizarFoto(int ID_FOTO)
        {
          string caminhoImagem = _RomaneioRepositorio.BuscarFoto(ID_FOTO);
           return Content($"<img src='{caminhoImagem}' class='img-fluid' alt='Imagem' />");
 
        }

        [HttpPost]
        public JsonResult ExcluirFoto(int ID_FOTO, int AUTONUM_ROMANEIO)
        {
            if (ID_FOTO != null)
            {
                string caminhoImagem = _RomaneioRepositorio.BuscarFoto(ID_FOTO);

                caminhoImagem = Path.Combine(Server.MapPath(caminhoImagem)) ;
                if (System.IO.File.Exists(caminhoImagem) == true)
                {
                    System.IO.File.Delete(caminhoImagem);
                }

                _RomaneioRepositorio.ExcluirFoto(ID_FOTO, AUTONUM_ROMANEIO);
                return Json(new { success = true });


            }
            return Json(new { success = false, message = "Item não encontrado" });
        }
    }
}