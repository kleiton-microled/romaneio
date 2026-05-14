using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;

namespace Romaneio.Repositorios
{
    public class ContainerReeferPatioRepositorio : IContainerReeferPatioRepositorio
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

        public CntrReeferBuscaFinalResult BuscarPorFinal(string final, int patioUsuario)
        {
            var r = new CntrReeferBuscaFinalResult();
            final = (final ?? string.Empty).Trim().ToUpperInvariant();
            if (final.Length == 0)
            {
                r.SUCESSO = false;
                r.MENSAGEM = "Informe o final do conteiner";
                return r;
            }

            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("PATIO", patioUsuario);
                var p = new DynamicParameters(filtro.Parametros);
                p.Add("Final", final);

                var sqlCount = @"
SELECT COUNT(1)
FROM operador..VW_INVENT_SISTEMAS
WHERE " + filtro.Sql + @"
  AND SUBSTRING(ID_CONTEINER, 8, 5) = @Final
  AND FLAG_REEFER = 'Y'
  AND FLAG_REEFER_DESLIGADO = 'N'";

                var n = con.QuerySingle<int>(sqlCount, p, commandTimeout: Config.QueryTimeoutInSeconds());
                if (n == 0)
                {
                    r.SUCESSO = false;
                    r.MENSAGEM = "Nenhum conteiner Reefer com este final";
                    return r;
                }
                if (n > 1)
                {
                    r.SUCESSO = false;
                    r.MENSAGEM = "Existe mais de 1 conteiner reefer com este final. Favor informar a identificacao completa";
                    return r;
                }

                var sqlId = @"
SELECT TOP 1 id_conteiner
FROM operador..VW_INVENT_SISTEMAS
WHERE " + filtro.Sql + @"
  AND SUBSTRING(ID_CONTEINER, 8, 5) = @Final
  AND FLAG_REEFER = 'Y'
  AND FLAG_REEFER_DESLIGADO = 'N'";

                r.ID_CONTEINER = (con.QueryFirstOrDefault<string>(sqlId, p, commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty).Trim();
                r.SUCESSO = true;
                return r;
            }
        }

        public ConteinerReeferDetalheDto CarregarConteiner(string idConteiner, int patioUsuario)
        {
            var dto = new ConteinerReeferDetalheDto();
            idConteiner = (idConteiner ?? string.Empty).Trim().ToUpperInvariant();
            dto.ID_CONTEINER = idConteiner;
            if (idConteiner.Length == 0)
            {
                dto.ERRO = "Conteiner nao informado";
                return dto;
            }

            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("PATIO", patioUsuario);
                var p = new DynamicParameters(filtro.Parametros);
                p.Add("Id", idConteiner);

                var sql = @"
SELECT
    autonum,
    data_ent_temp,
    sistema,
    temperature,
    scale,
    umidade,
    ventilacao,
    flag_reefer,
    flag_reefer_desligado
FROM operador..VW_INVENT_SISTEMAS
WHERE " + filtro.Sql + @"
  AND id_conteiner = @Id
  AND FLAG_REEFER = 'Y'
  AND FLAG_REEFER_DESLIGADO = 'N'";

                var row = con.Query(sql, p, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                if (row == null)
                {
                    dto.ERRO = "Conteiner nao encontrado";
                    return dto;
                }

                var d = (IDictionary<string, object>)row;
                var flagReefer = AsStr(Get(d, "flag_reefer", "FLAG_REEFER"));
                var flagDesl = AsStr(Get(d, "flag_reefer_desligado", "FLAG_REEFER_DESLIGADO"));
                if (flagReefer == "N")
                {
                    dto.ERRO = "Conteiner nao e reefer";
                    return dto;
                }
                if (flagDesl == "Y")
                {
                    dto.ERRO = "Conteiner reefer captado com carga DRY!";
                    return dto;
                }

                dto.AUTONUM = ToInt64(Get(d, "autonum", "AUTONUM"));
                dto.SISTEMA = AsStr(Get(d, "sistema", "SISTEMA"));
                dto.DATA_ENT_TEMP = FmtDataHora(Get(d, "data_ent_temp", "DATA_ENT_TEMP"));
                dto.TXT_SET_POINT = AsStr(Get(d, "temperature", "TEMPERATURE"));
                dto.TXT_ESCALA = AsStr(Get(d, "scale", "SCALE"));

                var umidadeDb = AsStr(Get(d, "umidade", "UMIDADE"));
                var ventDb = AsStr(Get(d, "ventilacao", "VENTILACAO"));
                dto.TXT_UMIDADE_SET = string.IsNullOrWhiteSpace(umidadeDb) ? "OFF" : umidadeDb;
                dto.TXT_VENTILACAO_SET = string.IsNullOrWhiteSpace(ventDb) ? "CLOSED" : ventDb;

                dto.AUTONUM_IPA = 0;
                dto.AUTONUM_RDX = 0;
                dto.AUTONUM_OP = 0;
                if (dto.SISTEMA == "I")
                    dto.AUTONUM_IPA = dto.AUTONUM;
                else if (dto.SISTEMA == "R")
                    dto.AUTONUM_RDX = dto.AUTONUM;
                else if (dto.SISTEMA == "O")
                    dto.AUTONUM_OP = dto.AUTONUM;

                dto.TXT_MIN = "";
                dto.TXT_MAX = "";
                dto.ALERTA_REEFER_DESLIGADO = false;
                dto.PLUG_OFF_TEXTO = "";

                if (dto.SISTEMA == "I" && dto.AUTONUM_IPA > 0)
                {
                    var bl = con.Query(
                        @"SELECT CONVERT(VARCHAR(10), dt_plug_off, 103) + ' ' + CONVERT(VARCHAR(8), dt_plug_off, 108) AS plug_fmt,
                                 tempmin, tempmax, dt_plug_off
                          FROM sgipa..TB_CNTR_BL WHERE AUTONUM = @A",
                        new { A = dto.AUTONUM_IPA },
                        commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    if (bl != null)
                    {
                        var bd = (IDictionary<string, object>)bl;
                        var po = AsStr(Get(bd, "plug_fmt", "PLUG_FMT"));
                        if (!string.IsNullOrWhiteSpace(po))
                        {
                            dto.ALERTA_REEFER_DESLIGADO = true;
                            dto.PLUG_OFF_TEXTO = "PLUG OFF:" + po;
                        }
                        dto.TXT_MIN = AsStr(Get(bd, "tempmin", "TEMPMIN"));
                        dto.TXT_MAX = AsStr(Get(bd, "tempmax", "TEMPMAX"));
                    }
                }

                return dto;
            }
        }

        public IReadOnlyList<ReeferHistoricoLinhaDto> ListarHistorico(long autonumIpa, long autonumRdx, long autonumOp)
        {
            if (autonumIpa <= 0 && autonumRdx <= 0 && autonumOp <= 0)
                return Array.Empty<ReeferHistoricoLinhaDto>();

            string filtroCol;
            object param;
            if (autonumIpa > 0)
            {
                filtroCol = "A.AUTONUM_CNTR_IPA = @Autonum";
                param = new { Autonum = autonumIpa };
            }
            else if (autonumRdx > 0)
            {
                filtroCol = "A.AUTONUM_CNTR_RDX = @Autonum";
                param = new { Autonum = autonumRdx };
            }
            else
            {
                filtroCol = "A.AUTONUM_CNTR_OP = @Autonum";
                param = new { Autonum = autonumOp };
            }

            var sql = @"
SELECT
    LTRIM(LEFT(CONVERT(CHAR(20), A.DT_MONITORING, 103), 5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), A.DT_MONITORING, 108), 5)) AS DT_MONITORING,
    A.TEMP_SUP + ' ' + ISNULL(A.ESCALA, '') AS TEMP_SUP,
    A.TEMP_RET + ' ' + ISNULL(A.ESCALA, '') AS TEMP_RET,
    A.UMIDADE,
    A.VENTILACAO,
    ISNULL(B.USUARIO, '') AS USUARIO,
    A.AUTONUM
FROM OPERADOR..TB_MONITORING A
LEFT JOIN SGIPA..TB_CAD_USUARIOS B ON A.USUARIO = B.AUTONUM
WHERE " + filtroCol + @"
ORDER BY A.AUTONUM DESC";

            using (var con = Open())
            {
                return con.Query<ReeferHistoricoLinhaDto>(sql, param, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public SalvarMonitoramentoResult SalvarMonitoramento(SalvarMonitoramentoRequest request, int patioUsuario, int usuarioId)
        {
            var res = new SalvarMonitoramentoResult();
            request = request ?? new SalvarMonitoramentoRequest();

            if (usuarioId <= 0)
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Usuario invalido na sessao";
                return res;
            }

            if (request.AUTONUM_IPA <= 0 && request.AUTONUM_RDX <= 0 && request.AUTONUM_OP <= 0)
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Conteiner invalido";
                return res;
            }

            var tempSup = ConsisteTemp(request.TXT_TEMP_SUP);
            var tempRet = ConsisteTemp(request.TXT_TEMP_RET);
            var umid = ConsisteUmidade(request.TXT_UMIDADE);
            var vent = ConsisteVentilacao(request.TXT_VENTILACAO);
            var escala = (request.TXT_ESCALA ?? string.Empty).Trim().ToUpperInvariant();

            if (string.IsNullOrEmpty(tempSup))
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Temperatura fornecida nao informada";
                return res;
            }
            if (string.IsNullOrEmpty(tempRet))
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Temperatura de retorno nao informada";
                return res;
            }
            if (string.IsNullOrEmpty(umid))
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Umidade nao informada";
                return res;
            }
            if (string.IsNullOrEmpty(vent))
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Ventilacao nao informada";
                return res;
            }
            if (string.IsNullOrEmpty(escala))
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Escala nao informada";
                return res;
            }
            if (escala != "F" && escala != "C")
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Escala invalida (use F ou C)";
                return res;
            }

            var setPoint = (request.TXT_SET_POINT ?? string.Empty).Trim().Replace(",", ".");
            decimal supVal;
            decimal setVal;
            if (decimal.TryParse(tempSup, NumberStyles.Any, CultureInfo.InvariantCulture, out supVal) &&
                decimal.TryParse(setPoint, NumberStyles.Any, CultureInfo.InvariantCulture, out setVal))
            {
                if (Math.Abs(supVal - setVal) > 2 && !request.CONFIRMAR_DIVERGENCIA_TEMPERATURA)
                {
                    res.SUCESSO = false;
                    res.EXIGE_CONFIRMACAO_DIVERGENCIA_TEMPERATURA = true;
                    res.MENSAGEM = "Atencao. Temperatura fornecida(Supply) muito divergente da temperatura de Set Point";
                    return res;
                }
            }

            var idCntr = (request.ID_CONTEINER ?? string.Empty).Trim().ToUpperInvariant();
            if (idCntr.Length == 0)
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Conteiner invalido";
                return res;
            }

            var patioCol = MapearPatioColunaTbMonitoring(patioUsuario);

            using (var con = Open())
            {
                con.Open();
                using (var tx = con.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        con.Execute(
                            @"INSERT INTO OPERADOR..TB_MONITORING (
                                DT_MONITORING, ID_CONTEINER,
                                AUTONUM_CNTR_IPA, AUTONUM_CNTR_RDX, AUTONUM_CNTR_OP,
                                TEMP_SETPOINT, UMIDADE_SETPOINT, VENTILACAO_SETPOINT,
                                TEMP_SUP, TEMP_RET, VENTILACAO, UMIDADE, ESCALA,
                                PATIO, USUARIO)
                              VALUES (
                                GETDATE(), @Id,
                                @Ipa, @Rdx, @Op,
                                @Sp, @Us, @Vs,
                                @Tsup, @Tret, @Vent, @Umid, @Esc,
                                @Patio, @Usuario)",
                            new
                            {
                                Id = idCntr,
                                Ipa = request.AUTONUM_IPA,
                                Rdx = request.AUTONUM_RDX,
                                Op = request.AUTONUM_OP,
                                Sp = request.TXT_SET_POINT ?? "",
                                Us = request.TXT_UMIDADE_SET ?? "",
                                Vs = request.TXT_VENTILACAO_SET ?? "",
                                Tsup = tempSup,
                                Tret = tempRet,
                                Vent = vent,
                                Umid = umid,
                                Esc = escala,
                                Patio = patioCol,
                                Usuario = usuarioId
                            },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        if (request.AUTONUM_IPA > 0)
                        {
                            con.Execute(
                                "UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF = NULL WHERE AUTONUM = @A",
                                new { A = request.AUTONUM_IPA },
                                tx,
                                Config.QueryTimeoutInSeconds());
                        }

                        if (request.AUTONUM_RDX > 0)
                        {
                            con.Execute(
                                "UPDATE REDEX..TB_PATIO SET DT_PLUG_OFF = NULL WHERE AUTONUM_PATIO = @A",
                                new { A = request.AUTONUM_RDX },
                                tx,
                                Config.QueryTimeoutInSeconds());
                        }

                        tx.Commit();
                        res.SUCESSO = true;
                        res.MENSAGEM = "Monitoramento registrado";
                        return res;
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        res.SUCESSO = false;
                        res.MENSAGEM = ex.Message;
                        return res;
                    }
                }
            }
        }

        private static int MapearPatioColunaTbMonitoring(int patioUsuario)
        {
            return patioUsuario == 7 ? 1 : patioUsuario;
        }

        public RegistrarPlugOffResult RegistrarPlugOff(long autonumIpa, int patioUsuario)
        {
            var r = new RegistrarPlugOffResult();
            if (autonumIpa <= 0)
            {
                r.SUCESSO = false;
                r.MENSAGEM = "Conteiner invalido";
                return r;
            }

            using (var con = Open())
            {
                var filtroV = MontarFiltroPatio("V.PATIO", patioUsuario);
                var p = new DynamicParameters(filtroV.Parametros);
                p.Add("Autonum", autonumIpa);

                var sqlPlug = @"
SELECT TOP 1 C.AUTONUM, C.DT_PLUG_OFF
FROM sgipa..TB_CNTR_BL C
INNER JOIN operador..VW_INVENT_SISTEMAS V ON V.autonum = C.AUTONUM AND V.sistema = 'I'
WHERE C.AUTONUM = @Autonum AND " + filtroV.Sql;

                var rowPlug = con.QueryFirstOrDefault<PlugOffVerificacaoRow>(sqlPlug, p, commandTimeout: Config.QueryTimeoutInSeconds());
                if (rowPlug == null)
                {
                    r.SUCESSO = false;
                    r.MENSAGEM = "Conteiner IPA nao localizado no patio";
                    return r;
                }
                if (rowPlug.DT_PLUG_OFF != null)
                {
                    r.SUCESSO = false;
                    r.MENSAGEM = "Operacao nao permitida: reefer desligado (plug off ja registrado)";
                    return r;
                }

                con.Execute(
                    "UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF = GETDATE() WHERE AUTONUM = @A",
                    new { A = autonumIpa },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                r.SUCESSO = true;
                r.MENSAGEM = "Plug OFF registrado";
                return r;
            }
        }

        public IReadOnlyList<ReeferEntradaPrevistaDto> ListarEntradasPrevistas(int patioUsuario)
        {
            var filtro = MontarFiltroPatio("C.PATIO", patioUsuario);
            var sql = @"
SELECT
    C.ID_CONTEINER,
    CONVERT(VARCHAR(20), C.TAMANHO) AS TAMANHO,
    C.TEMP + ' ' + ISNULL(C.ESCALA, '') AS SETPOINT,
    REPLACE(
        CASE WHEN V.DT_ATRACACAO IS NULL
            THEN LEFT(CONVERT(CHAR(20), V.DT_ATRACACAOP, 103), 10)
            ELSE LEFT(CONVERT(CHAR(20), V.DT_ATRACACAO, 103), 10)
        END, '01/01/1900', '') AS DT_ATRACACAO,
    CASE ISNULL(A.DESCR, '-') WHEN '-' THEN AP.DESCR ELSE A.DESCR END AS ARM_ATRACACAO,
    LTRIM(LEFT(CONVERT(CHAR(20), T.HR_ENTRADA_PREV, 103), 5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), T.HR_ENTRADA_PREV, 108), 5)) AS ENTRADA_PREVISTA
FROM SGIPA..TB_CNTR_BL C
INNER JOIN SGIPA..DTE_TB_VIAGENS V ON C.VIAGEM = V.VIAGEM
LEFT JOIN SGIPA..DTE_TB_ARMAZENS AP ON V.ARM_ATRACACAOP = AP.CODE
LEFT JOIN SGIPA..DTE_TB_ARMAZENS A ON V.ARM_ATRACACAO = A.CODE
LEFT JOIN SGIPA..TB_TTIMEUNT T ON C.ID_CONTEINER = T.SIGLA AND C.GMCI = T.GMCI
WHERE C.FLAG_TERMINAL = 0
  AND C.GMCI IS NOT NULL
  AND " + filtro.Sql + @"
  AND ISNULL(C.TEMP, ' ') <> ' '
  AND ISNULL(C.flag_desligado, 0) = 0
ORDER BY C.ID_CONTEINER";

            using (var con = Open())
            {
                return con.Query<ReeferEntradaPrevistaDto>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public IReadOnlyList<ReeferEstoqueLinhaDto> ListarEstoqueReefer(int patioUsuario, bool apenasDesligados, bool comAgendamentoSaida, bool posicionados)
        {
            var filtro = MontarFiltroPatio("c.Patio", patioUsuario);
            var sql = new System.Text.StringBuilder();
            sql.AppendLine(@"
SELECT
    C.ID_CONTEINER,
    LTRIM(LEFT(CONVERT(CHAR(20), C.DT_ENTRADA, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), C.DT_ENTRADA, 108), 5)) AS DT_ENTRADA,
    C.YARD,
    LTRIM(LEFT(CONVERT(CHAR(20), GD.PERIODO_INICIAL, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), GD.PERIODO_INICIAL, 108), 5)) AS PERIODO_INICIAL,
    C.TEMP + ' ' + ISNULL(C.ESCALA, '') AS SETPOINT,
    M1.TEMP_RET AS TEMP_RET_M1,
    LTRIM(LEFT(CONVERT(CHAR(20), M1.DT_MONITORING, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), M1.DT_MONITORING, 108), 5)) AS DT1,
    M2.TEMP_RET AS TEMP_RET_M2,
    LTRIM(LEFT(CONVERT(CHAR(20), M2.DT_MONITORING, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), M2.DT_MONITORING, 108), 5)) AS DT2,
    CASE WHEN c.dt_plug_off IS NOT NULL THEN 1 ELSE 0 END AS TEM_PLUG_OFF
FROM SGIPA..TB_CNTR_BL C
LEFT JOIN (
    SELECT MIN(AUTONUM) AS A_PRIM, AUTONUM_CNTR_IPA FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA
) PRIM ON C.AUTONUM = PRIM.AUTONUM_CNTR_IPA
LEFT JOIN OPERADOR..TB_MONITORING M1 ON PRIM.A_PRIM = M1.AUTONUM
LEFT JOIN (
    SELECT MAX(AUTONUM) AS A_ULT, AUTONUM_CNTR_IPA FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA
) ULT ON C.AUTONUM = ULT.AUTONUM_CNTR_IPA
LEFT JOIN OPERADOR..TB_MONITORING M2 ON ULT.A_ULT = M2.AUTONUM
LEFT JOIN OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA = GD.AUTONUM_GD_RESERVA
WHERE ");
            sql.Append(filtro.Sql);
            sql.AppendLine(@" AND C.FLAG_TERMINAL = 1 AND C.FLAG_HISTORICO = 0 AND C.DT_ENTRADA IS NOT NULL AND C.FLAG_DESLIGADO = 0
  AND ISNULL(C.TEMP, ' ') <> ' '");

            if (apenasDesligados)
                sql.AppendLine(" AND c.dt_plug_off IS NOT NULL");
            if (comAgendamentoSaida)
                sql.AppendLine(" AND GD.PERIODO_INICIAL IS NOT NULL");
            if (posicionados)
                sql.AppendLine(" AND (c.yard LIKE 'CF%' OR C.YARD LIKE 'CONF%')");

            sql.AppendLine(" ORDER BY C.ID_CONTEINER");

            using (var con = Open())
            {
                return con.Query<ReeferEstoqueLinhaDto>(sql.ToString(), filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public IReadOnlyList<ReeferDesligadoLinhaDto> ListarUnidadesDesligadas(int patioUsuario)
        {
            var filtro = MontarFiltroPatio("c.Patio", patioUsuario);
            var sql = @"
SELECT
    C.ID_CONTEINER,
    LTRIM(LEFT(CONVERT(CHAR(20), C.DT_ENTRADA, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), C.DT_ENTRADA, 108), 5)) AS DT_ENTRADA,
    C.YARD,
    LTRIM(LEFT(CONVERT(CHAR(20), GD.PERIODO_INICIAL, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), GD.PERIODO_INICIAL, 108), 5)) AS PERIODO_INICIAL,
    C.TEMP + ' ' + ISNULL(C.ESCALA, '') AS SETPOINT,
    M1.TEMP_RET AS TEMP_RET_M1,
    LTRIM(LEFT(CONVERT(CHAR(20), M1.DT_MONITORING, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), M1.DT_MONITORING, 108), 5)) AS DT1,
    M2.TEMP_RET AS TEMP_RET_M2,
    LTRIM(LEFT(CONVERT(CHAR(20), M2.DT_MONITORING, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), M2.DT_MONITORING, 108), 5)) AS DT2,
    CASE WHEN c.dt_plug_off IS NOT NULL THEN 1 ELSE 0 END AS TEM_PLUG_OFF,
    LTRIM(LEFT(CONVERT(CHAR(20), c.DT_plug_off, 103), 10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR(20), c.dt_plug_off, 108), 5)) AS DT_PLUG_OFF_FMT
FROM SGIPA..TB_CNTR_BL C
LEFT JOIN (
    SELECT MIN(AUTONUM) AS A_PRIM, AUTONUM_CNTR_IPA FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA
) PRIM ON C.AUTONUM = PRIM.AUTONUM_CNTR_IPA
LEFT JOIN OPERADOR..TB_MONITORING M1 ON PRIM.A_PRIM = M1.AUTONUM
LEFT JOIN (
    SELECT MAX(AUTONUM) AS A_ULT, AUTONUM_CNTR_IPA FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA
) ULT ON C.AUTONUM = ULT.AUTONUM_CNTR_IPA
LEFT JOIN OPERADOR..TB_MONITORING M2 ON ULT.A_ULT = M2.AUTONUM
LEFT JOIN OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA = GD.AUTONUM_GD_RESERVA
WHERE " + filtro.Sql + @"
  AND C.FLAG_TERMINAL = 1 AND C.FLAG_HISTORICO = 0 AND C.DT_ENTRADA IS NOT NULL AND C.FLAG_DESLIGADO = 0
  AND ISNULL(C.TEMP, ' ') <> ' '
  AND c.dt_plug_off IS NOT NULL
ORDER BY c.dt_plug_off DESC";

            using (var con = Open())
            {
                return con.Query<ReeferDesligadoLinhaDto>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        private sealed class PlugOffVerificacaoRow
        {
            public long AUTONUM { get; set; }
            public DateTime? DT_PLUG_OFF { get; set; }
        }

        #region Helpers (consistencia legado Microled + patio)

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

        private static string ConsisteTemp(string temperatura)
        {
            var t = (temperatura ?? string.Empty).Trim().ToUpperInvariant().Replace(",", ".").Replace("+", "");
            if (t.Length == 0) return "";
            decimal d;
            return decimal.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out d) ? t : "";
        }

        private static string ConsisteUmidade(string umidade)
        {
            var u = (umidade ?? string.Empty).Trim();
            if (u.Length == 0) return "OFF";
            decimal d;
            if (!decimal.TryParse(u.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return "OFF";
            if (d == 0) return "OFF";
            return d.ToString(CultureInfo.InvariantCulture);
        }

        private static string ConsisteVentilacao(string ventilacao)
        {
            var v = (ventilacao ?? string.Empty).Trim();
            if (v.Length == 0) return "CLOSED";
            decimal d;
            if (!decimal.TryParse(v.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return "CLOSED";
            if (d == 0) return "CLOSED";
            return d.ToString(CultureInfo.InvariantCulture);
        }

        private static object Get(IDictionary<string, object> d, params string[] keys)
        {
            foreach (var k in keys)
            {
                foreach (var kv in d)
                {
                    if (string.Equals(kv.Key, k, StringComparison.OrdinalIgnoreCase))
                        return kv.Value;
                }
            }
            return null;
        }

        private static string AsStr(object o)
        {
            if (o == null || o is DBNull) return "";
            return Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
        }

        private static long ToInt64(object o)
        {
            if (o == null || o is DBNull) return 0;
            try { return Convert.ToInt64(o, CultureInfo.InvariantCulture); }
            catch { return 0; }
        }

        private static string FmtDataHora(object o)
        {
            if (o == null || o is DBNull) return "";
            DateTime dt;
            if (DateTime.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
            return AsStr(o);
        }

        #endregion
    }
}
