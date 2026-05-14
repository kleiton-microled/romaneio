using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace Romaneio.Repositorios
{
    public class ManutencaoMarcantesArmazemRepositorio : IManutencaoMarcantesArmazemRepositorio
    {
        private class DivergenciaRow
        {
            public string embalagem { get; set; }

            public int? QUANTIDADE_REAL { get; set; }

            public int AUTONUM_CS { get; set; }
        }

        private class ListaRow
        {
            public int? ITEM { get; set; }

            public string DESCR_ARMAZEM { get; set; }

            public string POSICAO { get; set; }

            public int QTDE { get; set; }

            public int MARCANTE { get; set; }

            public int AUTONUMCS { get; set; }
        }

        public ManutencaoMarcantesArmazemFiltrarResult Filtrar(int lote)
        {
            var resultado = new ManutencaoMarcantesArmazemFiltrarResult();

            using (var con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                const string sqlDivergencia = @"
SELECT TOP 1
    embalagem,
    QUANTIDADE_REAL,
    AUTONUM_CS
FROM SGIPA..VW_WMS_DIVERGENCIA_QTDE
WHERE LOTE = @Lote";

                var div = con.QueryFirstOrDefault<DivergenciaRow>(
                    sqlDivergencia,
                    new { Lote = lote },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                if (div == null)
                {
                    resultado.MensagemErro = "Lote nao encontrado como divergente";
                    return resultado;
                }

                resultado.Embalagem = div.embalagem ?? string.Empty;
                resultado.QuantidadeReferencia = div.QUANTIDADE_REAL;
                resultado.AutonumCs = div.AUTONUM_CS;

                CarregarLinhas(con, resultado);
            }

            return resultado;
        }

        public ManutencaoMarcantesArmazemFiltrarResult Salvar(ManutencaoMarcantesArmazemSalvarRequest request)
        {
            var resultado = new ManutencaoMarcantesArmazemFiltrarResult
            {
                QuantidadeReferencia = request.QuantidadeReferencia,
                AutonumCs = request.AutonumCs,
                Embalagem = request.Embalagem ?? string.Empty
            };

            using (var con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();

                const string sqlUpdate = @"
UPDATE M
SET M.VOLUMES = @NovaQuantidade
FROM SGIPA..TB_MARCANTES M
INNER JOIN SGIPA..TB_CARGA_SOLTA C ON M.AUTONUM_CARGA = C.AUTONUM
WHERE M.AUTONUM = @MarcanteAutonum
  AND C.AUTONUM = @AutonumCs
  AND M.DT_ASSOCIACAO IS NOT NULL";

                var afetadas = con.Execute(
                    sqlUpdate,
                    new
                    {
                        NovaQuantidade = request.NovaQuantidade.Value,
                        MarcanteAutonum = request.MarcanteAutonum,
                        AutonumCs = request.AutonumCs
                    },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                if (afetadas == 0)
                {
                    resultado.MensagemErro = "Atualizacao nao permitida para este marcante e carga.";
                    return resultado;
                }

                AtualizarCabecalhoDivergenciaPorAutonumCs(con, resultado);
                CarregarLinhas(con, resultado);
            }

            return resultado;
        }

        /// <summary>
        /// Tenta preencher embalagem/referencia pela carga quando a view expuser AUTONUM_CS.
        /// Se a coluna nao existir, mantem valores ja informados no resultado.
        /// </summary>
        private static void AtualizarCabecalhoDivergenciaPorAutonumCs(SqlConnection con, ManutencaoMarcantesArmazemFiltrarResult resultado)
        {
            if (!resultado.AutonumCs.HasValue)
            {
                return;
            }

            const string sql = @"
SELECT TOP 1 embalagem, QUANTIDADE_REAL
FROM SGIPA..VW_WMS_DIVERGENCIA_QTDE
WHERE AUTONUM_CS = @AutonumCs";

            try
            {
                var row = con.QueryFirstOrDefault<DivergenciaRow>(
                    sql,
                    new { AutonumCs = resultado.AutonumCs.Value },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                if (row != null)
                {
                    resultado.Embalagem = row.embalagem ?? string.Empty;
                    if (row.QUANTIDADE_REAL.HasValue)
                    {
                        resultado.QuantidadeReferencia = row.QUANTIDADE_REAL;
                    }
                }
            }
            catch (SqlException)
            {
                // View pode nao filtrar por AUTONUM_CS; cabecalho permanece o enviado na requisicao.
            }
        }

        private static void CarregarLinhas(SqlConnection con, ManutencaoMarcantesArmazemFiltrarResult resultado)
        {
            if (!resultado.AutonumCs.HasValue)
            {
                return;
            }

            var autonumCs = resultado.AutonumCs.Value;

            var sqlLista = new StringBuilder();
            sqlLista.AppendLine("SELECT C.ITEM AS ITEM,");
            sqlLista.AppendLine("       I.DESCR_ARMAZEM AS DESCR_ARMAZEM,");
            sqlLista.AppendLine("       I.POSICAO AS POSICAO,");
            sqlLista.AppendLine("       M.VOLUMES AS QTDE,");
            sqlLista.AppendLine("       M.AUTONUM AS MARCANTE,");
            sqlLista.AppendLine("       M.AUTONUM_CARGA AS AUTONUMCS");
            sqlLista.AppendLine("FROM SGIPA..TB_MARCANTES M");
            sqlLista.AppendLine("INNER JOIN SGIPA..TB_CARGA_SOLTA C ON M.AUTONUM_CARGA = C.AUTONUM");
            sqlLista.AppendLine("LEFT JOIN (");
            sqlLista.AppendLine("    SELECT DESCR_ARMAZEM, POSICAO, MARCANTE");
            sqlLista.AppendLine("    FROM OPERADOR..VW_INVENT_ARMAZEM I");
            sqlLista.AppendLine("    WHERE AUTONUMCS = @AutonumCs");
            sqlLista.AppendLine("      AND QTDE > 0");
            sqlLista.AppendLine(") I ON M.AUTONUM = I.MARCANTE");
            sqlLista.AppendLine("WHERE C.AUTONUM = @AutonumCs");
            sqlLista.AppendLine("  AND M.DT_ASSOCIACAO IS NOT NULL");
            sqlLista.AppendLine("ORDER BY M.AUTONUM");

            var rows = con.Query<ListaRow>(
                sqlLista.ToString(),
                new { AutonumCs = autonumCs },
                commandTimeout: Config.QueryTimeoutInSeconds()).ToList();

            var indice = 1;
            foreach (var r in rows)
            {
                resultado.Linhas.Add(new ManutencaoMarcantesArmazemLinhaDto
                {
                    Indice = indice++,
                    DescrArmazem = r.DESCR_ARMAZEM ?? string.Empty,
                    Posicao = r.POSICAO ?? string.Empty,
                    Item = r.ITEM,
                    Qtde = r.QTDE,
                    Marcante = r.MARCANTE,
                    AutonumCs = r.AUTONUMCS
                });
            }

            resultado.TotalVolumesListados = resultado.Linhas.Sum(x => x.Qtde);
            if (resultado.QuantidadeReferencia.HasValue)
            {
                resultado.TotalConfereReferencia =
                    resultado.TotalVolumesListados == resultado.QuantidadeReferencia.Value;
            }
        }
    }
}
