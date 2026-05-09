using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace Romaneio.Repositorios
{
    public class LocalizacaoCarregamentosRepositorio : ILocalizacaoCarregamentosRepositorio
    {
        public string ObterDescricaoPatio(int patioUsuario)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var filtro = MontarFiltroPatio("autonum", patioUsuario);
                var sql = "SELECT TOP 1 descr_resumido FROM operador..tb_patios WHERE " + filtro.Sql;
                return con.QueryFirstOrDefault<string>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty;
            }
        }

        public LocalizacaoCarregamentosFiltroResult Filtrar(LocalizacaoCarregamentosFiltroRequest request, int patioUsuario)
        {
            request = request ?? new LocalizacaoCarregamentosFiltroRequest();
            var placa = (request.PLACA ?? string.Empty).Trim().ToUpperInvariant();

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var filtroCntr = MontarFiltroPatio("cc.patio", patioUsuario);
                var filtroLoc = MontarFiltroPatio("loc.patio", patioUsuario);
                var parametros = new DynamicParameters();

                if (filtroCntr.Parametros.ParameterNames.Contains("PATIO"))
                {
                    parametros.Add("PATIO", filtroCntr.Parametros.Get<int>("PATIO"));
                }

                parametros.Add("PLACA_FILTRO", placa);
                var filtroPlaca = string.IsNullOrWhiteSpace(placa)
                    ? string.Empty
                    : " AND (GN.PLACA = @PLACA_FILTRO OR GN.CARRETA = @PLACA_FILTRO) ";

                var sql = @"
SELECT
    gn.placa,
    gn.carreta,
    CONVERT(VARCHAR(5), o.dt_lib_ent_cam, 103) + ' ' + CONVERT(VARCHAR(5), o.dt_lib_ent_cam, 108) AS dt_lib_ent_cam,
    CONVERT(VARCHAR(5), gn.dt_gate_in, 103) + ' ' + CONVERT(VARCHAR(5), gn.dt_gate_in, 108) AS dt_gate_in,
    cc.id_conteiner AS carga,
    cc.yard AS quadra,
    ISNULL(rem.temp_rua, ' ') AS rua,
    t.fantasia AS nometransportadora,
    u.usuario,
    1 AS flag_cntr,
    cc.autonum AS autonum_cntr
FROM operador..tb_gate_new gn
INNER JOIN operador..tb_amr_gate amr ON gn.autonum = amr.gate
INNER JOIN sgipa..TB_CNTR_BL cc ON amr.CNTR_ipa = cc.AUTONUM
INNER JOIN operador..TB_CAD_TRANSPORTADORAS t ON gn.id_TRANSPORTADORA = t.AUTONUM
INNER JOIN sgipa..TB_ORDEM_CARREGAMENTO o ON amr.ID_OC = o.AUTONUM
LEFT JOIN operador..tb_remocao rem ON cc.autonum = rem.AUTONUM AND rem.sistema = 'I'
LEFT JOIN sgipa..tb_cad_usuarios u ON o.usu_lib_ent_cam = u.autonum
WHERE gn.dt_gate_out IS NULL
  AND gn.dt_gate_in > GETDATE() - 3
  AND " + filtroCntr.Sql + filtroPlaca + @"
UNION ALL
SELECT
    MAX(gn.placa) AS placa,
    MAX(gn.carreta) AS carreta,
    MAX(CONVERT(VARCHAR(5), loc.dt_lib_ent_cam, 103) + ' ' + CONVERT(VARCHAR(5), loc.dt_lib_ent_cam, 108)) AS dt_lib_ent_cam,
    MAX(CONVERT(VARCHAR(5), gn.dt_gate_in, 103) + ' ' + CONVERT(VARCHAR(5), gn.dt_gate_in, 108)) AS dt_gate_in,
    CONVERT(VARCHAR, SUM(loc.quantidade)) + ' vols ' AS carga,
    MAX(ISNULL(loc.armazem, ' ') + '-' + ISNULL(loc.yard, ' ')) AS quadra,
    ' ' AS rua,
    MAX(t.fantasia) AS nometransportadora,
    MAX(u.usuario) AS usuario,
    0 AS flag_cntr,
    0 AS autonum_cntr
FROM operador..tb_gate_new gn
INNER JOIN operador..tb_amr_gate amr ON gn.autonum = amr.gate
INNER JOIN operador..TB_CAD_TRANSPORTADORAS t ON gn.id_TRANSPORTADORA = t.AUTONUM
LEFT JOIN sgipa..VW_LOCAL_OC loc ON amr.ID_OC = loc.AUTONUMORDEM
LEFT JOIN sgipa..tb_cad_usuarios u ON loc.usu_lib_ent_cam = u.autonum
WHERE gn.dt_gate_out IS NULL
  AND gn.dt_gate_in > GETDATE() - 3
  AND " + filtroLoc.Sql + filtroPlaca + @"
GROUP BY gn.autonum
ORDER BY 1";

                var itens = con.Query<LocalizacaoCarregamentoItemDto>(
                    sql,
                    parametros,
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();

                return new LocalizacaoCarregamentosFiltroResult
                {
                    SUCESSO = true,
                    ITENS = itens,
                    TOTAL_ENTRADAS = itens.Count,
                    TOTAL_CNTR = itens.Count(i => i.FLAG_CNTR == 1),
                    TOTAL_CS = itens.Count(i => i.FLAG_CNTR != 1)
                };
            }
        }

        public MoverSaidaResult MoverParaSaida(MoverSaidaRequest request, int patioUsuario, int usuarioId)
        {
            if (request == null || request.AUTONUM_CNTR <= 0 || request.FLAG_CNTR != 1)
            {
                return FalhaSaida("Conteiner nao informado/invalido");
            }

            if (string.Equals((request.QUADRA_ATUAL ?? string.Empty).Trim(), "SAIDA", StringComparison.OrdinalIgnoreCase))
            {
                return FalhaSaida("Conteiner ja posicionado em SAIDA");
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                using (var tx = con.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        var filtroPatio = MontarFiltroPatio("a.patio", patioUsuario);
                        var sqlPosicao = @"
SELECT TOP 1
    ISNULL(v.valida, 0) AS valida,
    ISNULL(a.yard, '') AS yard_atual,
    a.autonum
FROM operador..VW_INVENT_SISTEMAS a
INNER JOIN operador..TB_YARD v ON a.PATIO = v.PATIO AND a.YARD = v.YARD
WHERE " + filtroPatio.Sql + " AND a.autonum = @AUTONUM_CNTR";

                        var parametrosPosicao = new DynamicParameters();
                        parametrosPosicao.Add("AUTONUM_CNTR", request.AUTONUM_CNTR);
                        if (filtroPatio.Parametros.ParameterNames.Contains("PATIO"))
                        {
                            parametrosPosicao.Add("PATIO", filtroPatio.Parametros.Get<int>("PATIO"));
                        }

                        var posicao = con.QueryFirstOrDefault<PosicaoInventarioDto>(
                            sqlPosicao,
                            parametrosPosicao,
                            tx,
                            Config.QueryTimeoutInSeconds());

                        if (posicao == null)
                        {
                            tx.Rollback();
                            return FalhaSaida("Conteiner nao localizado no patio");
                        }

                        var yardAtual = (posicao.YARD_ATUAL ?? string.Empty).Trim().ToUpperInvariant();
                        if (yardAtual == "SAIDA")
                        {
                            tx.Rollback();
                            return FalhaSaida("Conteiner ja posicionado em SAIDA");
                        }

                        if (posicao.VALIDA == 1 && ExisteConteinerAcima(con, tx, patioUsuario, yardAtual))
                        {
                            tx.Rollback();
                            return FalhaSaida("Existe um conteiner acima, impossivel movimentar");
                        }

                        con.Execute(
                            @"UPDATE SGIPA..TB_CNTR_BL
                              SET YARD = 'SAIDA'
                              WHERE AUTONUM = @AUTONUM_CNTR",
                            new { AUTONUM_CNTR = request.AUTONUM_CNTR },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        con.Execute(
                            @"INSERT INTO OPERADOR..TB_HIST_SHIFTING
                                  (CNTR, ORIGEM, DESTINO, DATA, TIPO, USUARIO, MOTIVO)
                              VALUES
                                  (@CNTR, @ORIGEM, 'SAIDA', GETDATE(), 'I', @USUARIO, 0)",
                            new
                            {
                                CNTR = request.AUTONUM_CNTR,
                                ORIGEM = yardAtual,
                                USUARIO = usuarioId
                            },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        tx.Commit();
                        return new MoverSaidaResult { SUCESSO = true, MENSAGEM = "Conteiner movimentado para SAIDA com sucesso" };
                    }
                    catch
                    {
                        tx.Rollback();
                        return FalhaSaida("Erro ao movimentar conteiner para SAIDA");
                    }
                }
            }
        }

        private bool ExisteConteinerAcima(SqlConnection con, IDbTransaction tx, int patioUsuario, string yardAtual)
        {
            var yardAcima = CalcularYardAcima(yardAtual);
            if (string.IsNullOrWhiteSpace(yardAcima))
            {
                return false;
            }

            var filtroPatio = MontarFiltroPatio("patio", patioUsuario);
            var sql = "SELECT COUNT(1) FROM operador..VW_INVENT_SISTEMAS WHERE " + filtroPatio.Sql + " AND YARD = @YARD";

            var parametros = new DynamicParameters();
            parametros.Add("YARD", yardAcima);
            if (filtroPatio.Parametros.ParameterNames.Contains("PATIO"))
            {
                parametros.Add("PATIO", filtroPatio.Parametros.Get<int>("PATIO"));
            }

            return con.QuerySingleOrDefault<int>(sql, parametros, tx, Config.QueryTimeoutInSeconds()) > 0;
        }

        private static string CalcularYardAcima(string yardAtual)
        {
            if (string.IsNullOrWhiteSpace(yardAtual))
            {
                return string.Empty;
            }

            var yard = yardAtual.Trim().ToUpperInvariant();
            if (yard.Length < 2)
            {
                return string.Empty;
            }

            var alturaTexto = yard.Substring(yard.Length - 1, 1);
            int altura;
            if (!int.TryParse(alturaTexto, out altura))
            {
                return string.Empty;
            }

            return yard.Substring(0, yard.Length - 1) + (altura + 1);
        }

        private static MoverSaidaResult FalhaSaida(string mensagem)
        {
            return new MoverSaidaResult
            {
                SUCESSO = false,
                MENSAGEM = mensagem
            };
        }

        private static FiltroPatioSql MontarFiltroPatio(string coluna, int patioUsuario)
        {
            var parametros = new DynamicParameters();
            if (patioUsuario == 1 || patioUsuario == 7)
            {
                return new FiltroPatioSql
                {
                    Sql = coluna + " IN (1,7)",
                    Parametros = parametros
                };
            }

            parametros.Add("PATIO", patioUsuario);
            return new FiltroPatioSql
            {
                Sql = coluna + " = @PATIO",
                Parametros = parametros
            };
        }

        private class FiltroPatioSql
        {
            public string Sql { get; set; }
            public DynamicParameters Parametros { get; set; }
        }

        private class PosicaoInventarioDto
        {
            public int VALIDA { get; set; }
            public string YARD_ATUAL { get; set; }
        }
    }
}
