using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Romaneio.Classes
{
  
    public class PaginatedList<T> : List<T>
    {
        public int PaginaAtual { get; private set; }
        public int TotalItens { get; private set; }
        public int TamanhoPagina { get; private set; }
        public int TotalPaginas { get; private set; }

        public PaginatedList(List<T> itens, int totalItens, int paginaAtual, int tamanhoPagina)
        {
            PaginaAtual = paginaAtual;
            TamanhoPagina = tamanhoPagina;
            TotalItens = totalItens;
            TotalPaginas = (int)Math.Ceiling(totalItens / (double)tamanhoPagina);

            this.AddRange(itens);
        }

        public static PaginatedList<T> Create(IQueryable<T> fonte, int paginaAtual, int tamanhoPagina)
        {
            var totalItens = fonte.Count();
            var itens = fonte.Skip((paginaAtual - 1) * tamanhoPagina).Take(tamanhoPagina).ToList();
            return new PaginatedList<T>(itens, totalItens, paginaAtual, tamanhoPagina);
        }
    }
}