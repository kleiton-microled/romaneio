using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;

namespace Romaneio.Repositorios
{
    public class DescargaDdcPatioRepositorio : IDescargaDdcPatioRepositorio
    {
        public string ObterDescricaoPatio(int patioUsuario)
        {
            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("autonum", patioUsuario);
                var sql = "SELECT TOP 1 descr_resumido FROM operador..tb_patios WHERE " + filtro.Sql;
                return con.QueryFirstOrDefault<string>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty;
            }
        }

        public IReadOnlyList<PlacaDdcListItemDto> ListarPlacasDdc(int patioUsuario)
        {
            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("cc.patio", patioUsuario);
                var sql = @"
SELECT DISTINCT
    a.autonum AS autonum_rs,
    RTRIM(LTRIM(oc.placa_c)) + ' (' + RTRIM(LTRIM(cc.id_conteiner)) + ')' AS placa
FROM sgipa..tb_registro_saida_cntr a
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
INNER JOIN sgipa..tb_ordem_carregamento oc ON a.ordem_carreg = oc.autonum
WHERE oc.flag_saida = 0
  AND oc.flag_ddc = 1
  AND " + filtro.Sql;

                return con.Query<PlacaDdcListItemDto>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public GateLotesResultDto ResolverGateELotes(long autonumRegistroSaida, int patioUsuario)
        {
            var r = new GateLotesResultDto { LOTES = new List<LoteAtivoDto>() };
            if (autonumRegistroSaida <= 0)
                return r;

            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("cc.patio", patioUsuario);
                var p = new DynamicParameters(filtro.Parametros);
                p.Add("Reg", autonumRegistroSaida);

                var sqlGate = @"
SELECT TOP 1 gn.autonum AS gate
FROM sgipa..tb_registro_saida_cntr a
INNER JOIN sgipa..tb_ordem_carregamento oc ON a.ordem_carreg = oc.autonum
INNER JOIN operador..tb_gate_new gn ON oc.placa_c = gn.placa
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
WHERE a.autonum = @Reg
  AND gn.flag_gate_in = 1
  AND ISNULL(gn.flag_gate_out, 0) = 0
  AND " + filtro.Sql;

                var gate = con.QueryFirstOrDefault<long?>(sqlGate, p, commandTimeout: Config.QueryTimeoutInSeconds());
                if (gate == null || gate.Value <= 0)
                    return r;

                r.GATE = gate.Value;

                var sqlLotes = @"
SELECT DISTINCT bl.autonum AS autonum_lote
FROM sgipa..tb_registro_saida_cntr a
INNER JOIN sgipa..tb_amr_cntr_bl c ON a.cntr = c.cntr
INNER JOIN sgipa..tb_bl bl ON c.bl = bl.autonum
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
WHERE a.autonum = @Reg
  AND bl.flag_ativo = 1
  AND " + filtro.Sql;

                r.LOTES = con.Query<LoteAtivoDto>(sqlLotes, p, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
                return r;
            }
        }

        public CarregarCabecalhoResultDto CarregarCabecalho(long autonumRegistroSaida, long autonumLote, int patioUsuario)
        {
            if (autonumRegistroSaida <= 0 || autonumLote <= 0)
                return null;

            var gateLotes = ResolverGateELotes(autonumRegistroSaida, patioUsuario);
            if (gateLotes.GATE <= 0)
                return null;
            if (gateLotes.LOTES == null || !gateLotes.LOTES.Any(x => x.AUTONUM_LOTE == autonumLote))
                return null;

            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("cc.patio", patioUsuario);
                var p = new DynamicParameters(filtro.Parametros);
                p.Add("Reg", autonumRegistroSaida);
                p.Add("Lote", autonumLote);

                var sql = @"
SELECT TOP 1
    a.cntr AS CNTR,
    @Gate AS GATE,
    a.ordem_carreg AS ID_OC,
    @Lote AS LOTE
FROM sgipa..tb_registro_saida_cntr a
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
WHERE a.autonum = @Reg
  AND " + filtro.Sql;

                p.Add("Gate", gateLotes.GATE);

                return con.QueryFirstOrDefault<CarregarCabecalhoResultDto>(sql, p, commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public IReadOnlyList<DescargaArmazemItemDto> ListarItens(long cntr, long gate, int patioUsuario)
        {
            if (cntr <= 0 || gate <= 0)
                return Array.Empty<DescargaArmazemItemDto>();

            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("cc.patio", patioUsuario);
                var p = new DynamicParameters(filtro.Parametros);
                p.Add("Cntr", cntr);
                p.Add("Gate", gate);

                var sql = @"
SELECT
    a.quantidade,
    c.descr AS descr_emb,
    ISNULL(a.finalizado, 0) AS finalizado,
    a.autonum,
    RTRIM(LTRIM(a.embalagem)) AS embalagem
FROM sgipa..tb_descarga_armazem a
LEFT JOIN operador..tb_amr_gate b ON a.autonum = b.id_descarga_armazem
INNER JOIN sgipa..dte_tb_embalagens c ON a.embalagem = c.code
LEFT JOIN operador..tb_gate_new d ON b.gate = d.autonum
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
WHERE a.cntr = @Cntr
  AND a.gate = @Gate
  AND " + filtro.Sql + @"
ORDER BY a.autonum";

                return con.Query<DescargaArmazemItemDto>(sql, p, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public IReadOnlyList<EmbalagemDto> ListarEmbalagens()
        {
            using (var con = Open())
            {
                const string sql = @"
SELECT RTRIM(LTRIM(code)) AS code, RTRIM(LTRIM(descr)) AS descr
FROM sgipa..dte_tb_embalagens
ORDER BY descr";
                return con.Query<EmbalagemDto>(sql, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public DescargaDdcPatioOperacaoResult SalvarItem(SalvarDescargaDdcRequest request, int patioUsuario, int usuarioId)
        {
            var res = new DescargaDdcPatioOperacaoResult();
            request = request ?? new SalvarDescargaDdcRequest();

            if (request.AutonumRegistroSaida <= 0)
            {
                res.MENSAGEM = "Registro de saida invalido";
                return res;
            }
            if (request.Cntr <= 0 || request.Gate <= 0 || request.Lote <= 0 || request.IdOc <= 0)
            {
                res.MENSAGEM = "Dados de cabecalho incompletos. Carregue a placa e o lote (F4).";
                return res;
            }
            if (request.Quantidade <= 0)
            {
                res.MENSAGEM = "Quantidade invalida";
                return res;
            }
            if (request.Finalizado != 0 && request.Finalizado != 1)
            {
                res.MENSAGEM = "Informe finalizacao de entrega";
                return res;
            }
            var finalizado = request.Finalizado == 1;

            var emb = (request.Embalagem ?? string.Empty).Trim();
            if (emb.Length == 0)
            {
                res.MENSAGEM = "Embalagem nao informada";
                return res;
            }

            using (var con = Open())
            {
                con.Open();

                if (!ValidarCabecalhoOperacional(con, request.AutonumRegistroSaida, request.Cntr, request.Gate, request.Lote, patioUsuario))
                {
                    res.MENSAGEM = "Dados inconsistentes com o registro de saida ou patio.";
                    return res;
                }

                using (var tx = con.BeginTransaction())
                {
                    try
                    {
                        if (request.AutonumItem <= 0)
                        {
                            const string sqlIns = @"
INSERT INTO sgipa..tb_descarga_armazem (cntr, quantidade, embalagem, usuario, lote, gate, id_oc, finalizado)
VALUES (@Cntr, @Qtd, @Emb, @Usu, @Lote, @Gate, @Oc, @Fin);
SELECT CAST(SCOPE_IDENTITY() AS BIGINT);";

                            var novoId = con.QuerySingle<long>(sqlIns, new
                            {
                                Cntr = request.Cntr,
                                Qtd = request.Quantidade,
                                Emb = emb,
                                Usu = usuarioId,
                                Lote = request.Lote,
                                Gate = request.Gate,
                                Oc = request.IdOc,
                                Fin = finalizado ? 1 : 0
                            }, tx, commandTimeout: Config.QueryTimeoutInSeconds());

                            if (novoId <= 0)
                            {
                                tx.Rollback();
                                res.MENSAGEM = "Falha ao inserir item de descarga.";
                                return res;
                            }
                        }
                        else
                        {
                            var saidaSum = con.QuerySingle<decimal>(@"
SELECT ISNULL(SUM(a.quantidade_saida), 0)
FROM operador..tb_amr_gate a
WHERE a.id_descarga_armazem = @Id",
                                new { Id = request.AutonumItem }, tx, commandTimeout: Config.QueryTimeoutInSeconds());

                            if (request.Quantidade < saidaSum)
                            {
                                tx.Rollback();
                                res.MENSAGEM = "Quantidade nao pode ser inferior a quantidade ja saida";
                                return res;
                            }

                            var embAtual = con.QueryFirstOrDefault<string>(@"
SELECT RTRIM(LTRIM(embalagem)) FROM sgipa..tb_descarga_armazem WHERE autonum = @Id AND cntr = @Cntr",
                                new { Id = request.AutonumItem, Cntr = request.Cntr }, tx, commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty;

                            var atualizaEmbalagem = saidaSum == 0 || string.Equals(emb, embAtual, StringComparison.OrdinalIgnoreCase);
                            if (!atualizaEmbalagem && saidaSum > 0)
                            {
                                res.AVISO = "Embalagem nao alterada pois ja existe saida para este item.";
                            }

                            if (atualizaEmbalagem)
                            {
                                const string sqlUpd = @"
UPDATE sgipa..tb_descarga_armazem
SET quantidade = @Qtd, finalizado = @Fin, embalagem = @Emb
WHERE autonum = @Id AND cntr = @Cntr AND gate = @Gate";
                                con.Execute(sqlUpd, new
                                {
                                    Qtd = request.Quantidade,
                                    Fin = finalizado ? 1 : 0,
                                    Emb = emb,
                                    Id = request.AutonumItem,
                                    Cntr = request.Cntr,
                                    Gate = request.Gate
                                }, tx, commandTimeout: Config.QueryTimeoutInSeconds());
                            }
                            else
                            {
                                const string sqlUpd2 = @"
UPDATE sgipa..tb_descarga_armazem
SET quantidade = @Qtd, finalizado = @Fin
WHERE autonum = @Id AND cntr = @Cntr AND gate = @Gate";
                                con.Execute(sqlUpd2, new
                                {
                                    Qtd = request.Quantidade,
                                    Fin = finalizado ? 1 : 0,
                                    Id = request.AutonumItem,
                                    Cntr = request.Cntr,
                                    Gate = request.Gate
                                }, tx, commandTimeout: Config.QueryTimeoutInSeconds());
                            }
                        }

                        AtualizarFlagDdcConteiner(con, tx, request.Cntr, finalizado);

                        tx.Commit();
                        res.SUCESSO = true;
                        res.MENSAGEM = string.IsNullOrEmpty(res.AVISO) ? "Registro gravado." : "Registro gravado. " + res.AVISO;
                        return res;
                    }
                    catch (Exception ex)
                    {
                        try { tx.Rollback(); } catch { /* ignore */ }
                        res.MENSAGEM = "Erro ao salvar: " + ex.Message;
                        return res;
                    }
                }
            }
        }

        public DescargaDdcPatioOperacaoResult ExcluirItem(long autonumItem, int patioUsuario)
        {
            var res = new DescargaDdcPatioOperacaoResult();
            if (autonumItem <= 0)
            {
                res.MENSAGEM = "Selecione um item para exclusao";
                return res;
            }

            using (var con = Open())
            {
                con.Open();
                using (var tx = con.BeginTransaction())
                {
                    try
                    {
                        var filtro = MontarFiltroPatio("cc.patio", patioUsuario);
                        var p = new DynamicParameters(filtro.Parametros);
                        p.Add("Id", autonumItem);

                        var cntrGate = con.QueryFirstOrDefault<CntrGateRow>(@"
SELECT a.cntr AS Cntr, a.gate AS Gate
FROM sgipa..tb_descarga_armazem a
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
WHERE a.autonum = @Id AND " + filtro.Sql, p, tx, commandTimeout: Config.QueryTimeoutInSeconds());

                        if (cntrGate == null || cntrGate.Cntr <= 0)
                        {
                            tx.Rollback();
                            res.MENSAGEM = "Item nao encontrado ou nao pertence ao patio.";
                            return res;
                        }

                        var saidaSum = con.QuerySingle<decimal>(@"
SELECT ISNULL(SUM(a.quantidade_saida), 0)
FROM operador..tb_amr_gate a
WHERE a.id_descarga_armazem = @Id",
                            new { Id = autonumItem }, tx, commandTimeout: Config.QueryTimeoutInSeconds());

                        if (saidaSum != 0)
                        {
                            tx.Rollback();
                            res.MENSAGEM = "Ja consta saida para este item - Exclusao cancelada";
                            return res;
                        }

                        con.Execute(@"DELETE FROM sgipa..tb_descarga_armazem WHERE autonum = @Id",
                            new { Id = autonumItem }, tx, commandTimeout: Config.QueryTimeoutInSeconds());

                        tx.Commit();
                        res.SUCESSO = true;
                        res.MENSAGEM = "Registro excluido!";
                        return res;
                    }
                    catch (Exception ex)
                    {
                        try { tx.Rollback(); } catch { }
                        res.MENSAGEM = "Erro ao excluir: " + ex.Message;
                        return res;
                    }
                }
            }
        }

        private static void AtualizarFlagDdcConteiner(SqlConnection con, IDbTransaction tx, long cntr, bool finalizado)
        {
            var flagAtual = con.QuerySingle<int>(@"
SELECT ISNULL(flag_ddc, 0) FROM sgipa..tb_cntr_bl WHERE autonum = @Cntr",
                new { Cntr = cntr }, tx, commandTimeout: Config.QueryTimeoutInSeconds());

            if (flagAtual == 0)
            {
                con.Execute(@"
UPDATE sgipa..tb_cntr_bl
SET flag_ddc = 1, inicio_ddc = GETDATE()
WHERE autonum = @Cntr AND ISNULL(flag_ddc, 0) = 0",
                    new { Cntr = cntr }, tx, commandTimeout: Config.QueryTimeoutInSeconds());
            }

            if (finalizado)
            {
                con.Execute(@"
UPDATE sgipa..tb_cntr_bl
SET flag_ddc = 2
WHERE autonum = @Cntr",
                    new { Cntr = cntr }, tx, commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        private static bool ValidarCabecalhoOperacional(SqlConnection con, long autonumRegistroSaida, long cntr, long gate, long lote, int patioUsuario)
        {
            var filtro = MontarFiltroPatio("cc.patio", patioUsuario);
            var p = new DynamicParameters(filtro.Parametros);
            p.Add("Reg", autonumRegistroSaida);
            p.Add("Cntr", cntr);
            p.Add("Gate", gate);
            p.Add("Lote", lote);

            var n = con.QuerySingle<int>(@"
SELECT COUNT(1)
FROM sgipa..tb_registro_saida_cntr a
INNER JOIN sgipa..tb_cntr_bl cc ON a.cntr = cc.autonum
INNER JOIN sgipa..tb_ordem_carregamento oc ON a.ordem_carreg = oc.autonum
INNER JOIN operador..tb_gate_new gn ON oc.placa_c = gn.placa AND gn.autonum = @Gate
WHERE a.autonum = @Reg
  AND a.cntr = @Cntr
  AND gn.flag_gate_in = 1
  AND ISNULL(gn.flag_gate_out, 0) = 0
  AND EXISTS (
      SELECT 1
      FROM sgipa..tb_amr_cntr_bl c
      INNER JOIN sgipa..tb_bl bl ON c.bl = bl.autonum
      WHERE c.cntr = a.cntr AND bl.autonum = @Lote AND bl.flag_ativo = 1
  )
  AND " + filtro.Sql, p, commandTimeout: Config.QueryTimeoutInSeconds());

            return n > 0;
        }

        private static SqlConnection Open() => new SqlConnection(Config.StringConexao());

        private sealed class FiltroPatioSql
        {
            public string Sql { get; set; }
            public DynamicParameters Parametros { get; set; }
        }

        private static FiltroPatioSql MontarFiltroPatio(string coluna, int patioUsuario)
        {
            var parametros = new DynamicParameters();
            if (patioUsuario == 1 || patioUsuario == 7)
            {
                return new FiltroPatioSql { Sql = coluna + " IN (1,7)", Parametros = parametros };
            }
            parametros.Add("PATIO", patioUsuario);
            return new FiltroPatioSql { Sql = coluna + " = @PATIO", Parametros = parametros };
        }

        private sealed class CntrGateRow
        {
            public long Cntr { get; set; }
            public long Gate { get; set; }
        }
    }
}
