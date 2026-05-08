using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace Romaneio.Repositorios
{
    public class AssociacaoMarcantesRepositorio : IAssociacaoMarcantesRepositorio
    {
        public List<string> ObterConteinersPorLote(int? lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(" SELECT DISTINCT ID_CONTEINER ");
                sb.Append(" FROM SGIPA..VW_WMS_SEM_MARCANTE ");
                sb.Append(" WHERE ID_CONTEINER IS NOT NULL ");

                if (lote.HasValue && lote.Value > 0)
                {
                    sb.Append(" AND LOTE = @LOTE ");
                }

                sb.Append(" ORDER BY ID_CONTEINER ");

                return con.Query<string>(sb.ToString(), new { LOTE = lote }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public AssociacaoMarcantesFiltroResult Filtrar(int? lote, string cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var itensNormal = CarregarItens(con, false, lote, cntr);
                if (itensNormal.Any())
                {
                    return new AssociacaoMarcantesFiltroResult
                    {
                        MODO_EXP = false,
                        ITENS = itensNormal
                    };
                }

                var itensExp = CarregarItens(con, true, lote, cntr);
                return new AssociacaoMarcantesFiltroResult
                {
                    MODO_EXP = itensExp.Any(),
                    ITENS = itensExp
                };
            }
        }

        public List<AssociacaoMarcanteTempDto> CarregarTemporarios(bool modoExp, int autonumRef)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var sql = modoExp
                    ? "SELECT MARCANTE, QTDE FROM SGIPA..TB_TEMP_MARCANTE_CEXP WHERE AUTONUM_CEXP = @AUTONUM_REF ORDER BY MARCANTE"
                    : "SELECT MARCANTE, QTDE FROM SGIPA..TB_TEMP_MARCANTE WHERE AUTONUM_CS = @AUTONUM_REF ORDER BY MARCANTE";

                return con.Query<AssociacaoMarcanteTempDto>(sql, new { AUTONUM_REF = autonumRef }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public string AdicionarTemp(AssociacaoMarcantesAdicionarRequest request)
        {
            if (request.MARCANTE <= 0)
            {
                return "Informe o marcante";
            }

            if (request.QTDE <= 0)
            {
                return "Informe a quantidade associada ao marcante";
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();

                var sqlExisteTemp = request.MODO_EXP
                    ? "SELECT 1 FROM SGIPA..TB_TEMP_MARCANTE_CEXP WHERE AUTONUM_CEXP = @AUTONUM_REF AND MARCANTE = @MARCANTE"
                    : "SELECT 1 FROM SGIPA..TB_TEMP_MARCANTE WHERE AUTONUM_CS = @AUTONUM_REF AND MARCANTE = @MARCANTE";

                var jaAdicionado = con.QueryFirstOrDefault<int?>(sqlExisteTemp, new { request.AUTONUM_REF, request.MARCANTE }, commandTimeout: Config.QueryTimeoutInSeconds());
                if (jaAdicionado.HasValue)
                {
                    return "Marcante ja adicionado";
                }

                var marcante = con.QueryFirstOrDefault<MarcanteStatusDto>(
                    "SELECT AUTONUM, DT_ASSOCIACAO, AUTONUM_CEXP FROM SGIPA..TB_MARCANTES WHERE AUTONUM = @MARCANTE",
                    new { request.MARCANTE },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                if (marcante == null)
                {
                    return "Nr de marcante invalido";
                }

                if (marcante.DT_ASSOCIACAO.HasValue)
                {
                    return "Nr de marcante ja utilizado";
                }

                if (marcante.AUTONUM_CEXP > 0)
                {
                    if (!request.MODO_EXP)
                    {
                        return "Nr de marcante de carga de exportacao";
                    }

                    if (marcante.AUTONUM_CEXP != request.AUTONUM_REF)
                    {
                        return "Marcante divergente ao lote exp.";
                    }
                }

                var sqlInsert = request.MODO_EXP
                    ? "INSERT INTO SGIPA..TB_TEMP_MARCANTE_CEXP(AUTONUM_CEXP, MARCANTE, QTDE) VALUES (@AUTONUM_REF, @MARCANTE, @QTDE)"
                    : "INSERT INTO SGIPA..TB_TEMP_MARCANTE(AUTONUM_CS, MARCANTE, QTDE) VALUES (@AUTONUM_REF, @MARCANTE, @QTDE)";

                con.Execute(sqlInsert, new { request.AUTONUM_REF, request.MARCANTE, request.QTDE }, commandTimeout: Config.QueryTimeoutInSeconds());
                return string.Empty;
            }
        }

        public string RemoverTemp(AssociacaoMarcantesRemoverRequest request)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var sqlDelete = request.MODO_EXP
                    ? "DELETE FROM SGIPA..TB_TEMP_MARCANTE_CEXP WHERE AUTONUM_CEXP = @AUTONUM_REF AND MARCANTE = @MARCANTE"
                    : "DELETE FROM SGIPA..TB_TEMP_MARCANTE WHERE AUTONUM_CS = @AUTONUM_REF AND MARCANTE = @MARCANTE";

                con.Execute(sqlDelete, new { request.AUTONUM_REF, request.MARCANTE }, commandTimeout: Config.QueryTimeoutInSeconds());
            }

            return string.Empty;
        }

        public string Salvar(AssociacaoMarcantesSalvarRequest request)
        {
            if (request.ITEM_SELECIONADO == null || request.ITEM_SELECIONADO.AUTONUMCS <= 0)
            {
                return "Carga nao selecionada";
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                using (var transaction = con.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        var temporarios = CarregarTemporariosInterno(con, transaction, request.MODO_EXP, request.ITEM_SELECIONADO.AUTONUMCS);
                        var somaTemp = temporarios.Sum(x => x.QTDE);

                        if (somaTemp != request.ITEM_SELECIONADO.QUANTIDADE)
                        {
                            transaction.Rollback();
                            return "Quantidades divergentes";
                        }

                        var qtdeCs = request.MODO_EXP
                            ? con.QuerySingleOrDefault<int>(
                                "SELECT ISNULL(QUANTIDADE - ISNULL(QUANTIDADE_SAIDA,0),0) FROM SGIPA..TB_CARGA_EXPORTACAO WHERE AUTONUM = @AUTONUMCS",
                                new { AUTONUMCS = request.ITEM_SELECIONADO.AUTONUMCS },
                                transaction,
                                Config.QueryTimeoutInSeconds())
                            : con.QuerySingleOrDefault<int>(
                                "SELECT ISNULL(QUANTIDADE_REAL,0) FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM = @AUTONUMCS",
                                new { AUTONUMCS = request.ITEM_SELECIONADO.AUTONUMCS },
                                transaction,
                                Config.QueryTimeoutInSeconds());

                        foreach (var item in temporarios)
                        {
                            var qtdeJaMarcada = request.MODO_EXP
                                ? con.QuerySingleOrDefault<int>(
                                    "SELECT ISNULL(SUM(VOLUMES),0) FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CEXP = @AUTONUMCS",
                                    new { AUTONUMCS = request.ITEM_SELECIONADO.AUTONUMCS },
                                    transaction,
                                    Config.QueryTimeoutInSeconds())
                                : con.QuerySingleOrDefault<int>(
                                    "SELECT ISNULL(SUM(VOLUMES),0) FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CARGA = @AUTONUMCS",
                                    new { AUTONUMCS = request.ITEM_SELECIONADO.AUTONUMCS },
                                    transaction,
                                    Config.QueryTimeoutInSeconds());

                            if (qtdeJaMarcada + item.QTDE <= qtdeCs)
                            {
                                if (request.MODO_EXP)
                                {
                                    con.Execute(
                                        "UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO = GETDATE(), AUTONUM_CEXP = @AUTONUMCS, VOLUMES = @QTDE WHERE AUTONUM = @MARCANTE",
                                        new
                                        {
                                            AUTONUMCS = request.ITEM_SELECIONADO.AUTONUMCS,
                                            QTDE = item.QTDE,
                                            MARCANTE = item.MARCANTE
                                        },
                                        transaction,
                                        Config.QueryTimeoutInSeconds());
                                }
                                else
                                {
                                    con.Execute(
                                        "UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO = GETDATE(), AUTONUM_CARGA = @AUTONUMCS, VOLUMES = @QTDE, BL = @BL, ITEM = @ITEM, CNTR = @CNTR WHERE AUTONUM = @MARCANTE",
                                        new
                                        {
                                            AUTONUMCS = request.ITEM_SELECIONADO.AUTONUMCS,
                                            QTDE = item.QTDE,
                                            BL = request.ITEM_SELECIONADO.LOTE,
                                            ITEM = request.ITEM_SELECIONADO.ITEM,
                                            CNTR = request.ITEM_SELECIONADO.AUTONUMCNTR,
                                            MARCANTE = item.MARCANTE
                                        },
                                        transaction,
                                        Config.QueryTimeoutInSeconds());
                                }
                            }
                        }

                        transaction.Commit();
                        return "MARCANTES ASSOCIADOS COM SUCESSO";
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return "Erro ao salvar associacao: " + ex.Message;
                    }
                }
            }
        }

        private List<AssociacaoMarcantesLinhaDto> CarregarItens(SqlConnection con, bool modoExp, int? lote, string cntr)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT LOTE, ITEM, QUANTIDADE, EMBALAGEM, ID_CONTEINER, MERCADORIA, ");
            sb.Append(modoExp ? "AUTONUM_CEXP AS AUTONUMCS, AUTONUMCNTR " : "AUTONUMCS, AUTONUMCNTR ");
            sb.Append("FROM ");
            sb.Append(modoExp ? "SGIPA..VW_WMS_SEM_MARCANTE_CEXP " : "SGIPA..VW_WMS_SEM_MARCANTE ");
            sb.Append("WHERE QUANTIDADE > 0 ");

            if (!lote.HasValue && string.IsNullOrWhiteSpace(cntr))
            {
                sb.Append("AND FLAG_HISTORICO = 0 ");
            }
            else
            {
                if (lote.HasValue && lote.Value > 0)
                {
                    sb.Append("AND LOTE = @LOTE ");
                }

                if (!string.IsNullOrWhiteSpace(cntr))
                {
                    sb.Append("AND ID_CONTEINER = @CNTR ");
                }
            }

            sb.Append("ORDER BY LOTE");

            var itens = con.Query<AssociacaoMarcantesLinhaDto>(sb.ToString(), new { LOTE = lote, CNTR = cntr }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();

            for (var i = 0; i < itens.Count; i++)
            {
                itens[i].INDICE = i + 1;
            }

            return itens;
        }

        private List<AssociacaoMarcanteTempDto> CarregarTemporariosInterno(SqlConnection con, IDbTransaction transaction, bool modoExp, int autonumRef)
        {
            var sql = modoExp
                ? "SELECT MARCANTE, QTDE FROM SGIPA..TB_TEMP_MARCANTE_CEXP WHERE AUTONUM_CEXP = @AUTONUM_REF ORDER BY MARCANTE"
                : "SELECT MARCANTE, QTDE FROM SGIPA..TB_TEMP_MARCANTE WHERE AUTONUM_CS = @AUTONUM_REF ORDER BY MARCANTE";

            return con.Query<AssociacaoMarcanteTempDto>(sql, new { AUTONUM_REF = autonumRef }, transaction, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
        }

        private class MarcanteStatusDto
        {
            public int AUTONUM { get; set; }
            public DateTime? DT_ASSOCIACAO { get; set; }
            public int AUTONUM_CEXP { get; set; }
        }
    }
}
