using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;

namespace Romaneio.Repositorios
{
    public class MovimentacaoContainerPatioRepositorio : IMovimentacaoContainerPatioRepositorio
    {
        #region Patio / cadastros

        public string ObterDescricaoPatio(int patioUsuario)
        {
            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("autonum", patioUsuario);
                var sql = "SELECT TOP 1 descr_resumido FROM operador..tb_patios WHERE " + filtro.Sql;
                return con.QueryFirstOrDefault<string>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty;
            }
        }

        public MovimentacaoContainerPatioPermissoesDto ObterPermissoesAvarias(int usuarioId)
        {
            var p = new MovimentacaoContainerPatioPermissoesDto();
            using (var con = Open())
            {
                p.ABA_AVARIAS_HABILITADA = ContarPermissao(con, usuarioId, 998, 5) > 0;
                p.PODE_EXCLUIR_AVARIA = ContarPermissao(con, usuarioId, 998, 3) > 0;
                p.PODE_INCLUIR_AVARIA = ContarPermissao(con, usuarioId, 998, 1) > 0;
                p.PODE_FINALIZAR_AVARIA = ContarPermissao(con, usuarioId, 999, 1) > 0;
            }
            return p;
        }

        private static int ContarPermissao(IDbConnection con, int usuarioId, int codFunc, int codTipoPerm)
        {
            var sql = @"
SELECT COUNT(1) CONTAR
FROM sgipa..TB_SYS_FUNCOES f
INNER JOIN sgipa..TB_SYS_GRP_PERMISSOES gp ON f.codfunc = gp.codfunc
INNER JOIN sgipa..TB_SYS_USER_GRUPOS ug ON gp.codgrupo = ug.codgrupo
INNER JOIN sgipa..TB_CAD_USUARIOS u ON ug.autonumuser = u.autonum
WHERE u.AUTONUM = @UsuarioId
  AND f.CODFUNC = @CodFunc
  AND gp.CODTIPOPERM = @CodTipoPerm";
            return con.QuerySingle<int>(sql, new { UsuarioId = usuarioId, CodFunc = codFunc, CodTipoPerm = codTipoPerm }, commandTimeout: Config.QueryTimeoutInSeconds());
        }

        public List<MotivoMovimentacaoDto> ListarMotivosAtivos()
        {
            using (var con = Open())
            {
                var rows = con.Query(
                    @"SELECT autonum AS AUTONUM, descricao AS DESCRICAO,
                    ISNULL(flag_camera,0) AS FLAG_CAMERA
                    FROM operador..tb_cad_motivo WHERE flag_ativo = 1 ORDER BY descricao",
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
                var list = new List<MotivoMovimentacaoDto>();
                foreach (var row in rows)
                {
                    var d = (IDictionary<string, object>)row;
                    list.Add(new MotivoMovimentacaoDto
                    {
                        AUTONUM = ToInt32Safe(GetCI(d, "AUTONUM", "autonum")),
                        DESCRICAO = AsStr(GetCI(d, "DESCRICAO", "descricao")),
                        FLAG_CAMERA = ToInt32Safe(GetCI(d, "FLAG_CAMERA", "flag_camera")) != 0
                    });
                }
                return list;
            }
        }

        public List<CameraPatioDto> ListarCameras(int patioUsuario)
        {
            using (var con = Open())
            {
                var filtro = MontarFiltroPatio("PATIO", patioUsuario);
                var sql = "SELECT autonum AS AUTONUM, descr AS DESCR FROM operador..TB_CAMERAS WHERE " + filtro.Sql + " ORDER BY descr";
                return con.Query<CameraPatioDto>(sql, filtro.Parametros, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public bool MotivoExigeCamera(int motivoAutonum)
        {
            using (var con = Open())
            {
                var n = con.QuerySingle<int>(
                    "SELECT COUNT(*) FROM operador..tb_cad_motivo WHERE autonum=@M AND ISNULL(flag_camera,0)=1",
                    new { M = motivoAutonum },
                    commandTimeout: Config.QueryTimeoutInSeconds());
                return n > 0;
            }
        }

        public int? ObterCameraPadraoPorYard(int patioUsuario, string yardDestino)
        {
            if (string.IsNullOrWhiteSpace(yardDestino)) return null;
            var filtro = MontarFiltroPatio("PATIO", patioUsuario);
            var sql = @"SELECT ISNULL(AUTONUM_CAMERA,0) FROM operador..TB_YARD WHERE (" + filtro.Sql +
                      @") AND (YARD=@Y OR YARD_EQ=@Y)";
            var p = new DynamicParameters(filtro.Parametros);
            p.Add("Y", yardDestino.Trim());
            using (var con = Open())
            {
                var v = con.QuerySingleOrDefault<int>(sql, p, commandTimeout: Config.QueryTimeoutInSeconds());
                return v > 0 ? v : (int?)null;
            }
        }

        #endregion

        #region Busca / detalhe

        public CntrPatioBuscaFinalResult BuscarPorFinal(string final, int patioUsuario)
        {
            var r = new CntrPatioBuscaFinalResult();
            final = (final ?? string.Empty).Trim().ToUpperInvariant();
            if (final.Length == 0)
            {
                r.SUCESSO = false;
                r.MENSAGEM = "Informe o final do conteiner";
                return r;
            }

            using (var con = Open())
            {
                var rows = con.Query(
                    "SELECT * FROM operador.dbo.fnInvent_Sistemas_YardF(@F, @P)",
                    new { F = final, P = patioUsuario },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();

                if (rows.Count == 0)
                {
                    r.SUCESSO = false;
                    r.MENSAGEM = "Nenhum conteiner com este final";
                    return r;
                }
                if (rows.Count > 1)
                {
                    r.SUCESSO = false;
                    r.MENSAGEM = "Existe mais de 1 conteiner com este final. Favor informar a identificacao completa";
                    return r;
                }

                var dr = (IDictionary<string, object>)rows[0];
                r.ID_CONTEINER = AsStr(GetCI(dr, "Id_Conteiner", "ID_CONTEINER", "id_conteiner"));
                r.SUCESSO = true;
                return r;
            }
        }

        public CntrPatioDetalheDto CarregarDetalhe(string idConteiner, int patioUsuario, int usuarioId)
        {
            idConteiner = (idConteiner ?? string.Empty).Trim().ToUpperInvariant();
            var dto = new CntrPatioDetalheDto { ID_CONTEINER = idConteiner };

            using (var con = Open())
            {
                var row = con.Query(
                    "SELECT * FROM operador.dbo.fnInvent_Sistemas_Col(@Id, @P)",
                    new { Id = idConteiner, P = patioUsuario },
                    commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                if (row == null)
                {
                    dto.TEXTO_CONTEXTUAL = "__ERRO__|Unidade nao encontrada no estoque";
                    return dto;
                }

                var d = (IDictionary<string, object>)row;
                dto.AUTONUM = Convert.ToInt64(GetCI(d, "autonum", "AUTONUM"), CultureInfo.InvariantCulture);
                dto.SISTEMA = AsStr(GetCI(d, "Sistema", "sistema"));
                dto.EF = AsStr(GetCI(d, "EF", "ef"));
                dto.TXT_GATE_IN = FmtData(GetCI(d, "data_ent_temp", "DATA_ENT_TEMP"));
                dto.NAVIO_VIAGEM = TrimJoin(AsStr(GetCI(d, "Nome", "NOME")), AsStr(GetCI(d, "Viagem", "VIAGEM")));
                dto.POD = AsStr(GetCI(d, "pod", "POD"));
                dto.BRUTO = AsStr(GetCI(d, "Bruto", "BRUTO"));
                dto.LACRES_RESUMO = string.Join(" ", new[]
                {
                    AsStr(GetCI(d, "Lacre1", "LACRE1")), AsStr(GetCI(d, "Lacre2", "LACRE2")), AsStr(GetCI(d, "Lacre3", "LACRE3")),
                    AsStr(GetCI(d, "Lacre4", "LACRE4")), AsStr(GetCI(d, "lacre5", "LACRE5")), AsStr(GetCI(d, "lacre6", "LACRE6")), AsStr(GetCI(d, "lacre7", "LACRE7"))
                }.Where(x => !string.IsNullOrWhiteSpace(x)));
                dto.IMO = AsStr(GetCI(d, "IMO1", "imo1"));
                dto.TAMANHO = ToInt32Safe(GetCI(d, "Tamanho", "TAMANHO"));
                dto.TIPO_BASICO = AsStr(GetCI(d, "TipoBasico", "TIPOBASICO"));
                dto.YARD_ATUAL = AsStr(GetCI(d, "Yard", "YARD"));
                dto.GMCI = AsStr(GetCI(d, "GMCI", "gmci"));
                dto.FLAG_CNTR_BLOQ_CHAO_PATIO = ToInt32Safe(GetCI(d, "FLAG_CNTR_BLOQ_CHAO_PATIO", "flag_cntr_bloq_chao_patio"));

                var frd = AsStr(GetCI(d, "flag_reefer_desligado", "FLAG_REEFER_DESLIGADO"));
                var fr = AsStr(GetCI(d, "flag_reefer", "FLAG_REEFER"));
                dto.FLAG_REEFER_DESLIGADO_VISIVEL = frd == "Y";
                dto.FLAG_REEFER_LIGADO_VISIVEL = frd == "N" && fr == "Y";

                dto.LBL_DTA_VISIVEL = AsStr(GetCI(d, "flag_DTA", "FLAG_DTA")) == "Y";

                var fdes = ToInt32Safe(GetCI(d, "flag_DESTINADO_DESOVA", "FLAG_DESTINADO_DESOVA"));
                var regime = AsStr(GetCI(d, "Regime", "REGIME"));
                dto.LBL_DESOVA_VISIVEL = fdes == 1 || regime == "HP" || regime == "PP";
                dto.LBL_DESOVA_TEXTO = "DESTINADO DESOVA";

                var descrSig = AsStr(GetCI(d, "DESCRICAO_SIGVIG", "descricao_sigvig"));
                if (descrSig.IndexOf("SELECIONADO", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    descrSig.IndexOf("NO", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    if (fdes == 1 || regime == "HP" || regime == "PP")
                    {
                        dto.LBL_DESOVA_TEXTO = "DEST.DESOVA/MAPA";
                    }
                    else
                    {
                        dto.LBL_DESOVA_TEXTO = "SELECIONADO MAPA";
                    }
                    dto.LBL_DESOVA_VISIVEL = true;
                }

                dto.TARA = ToDecimalSafe(GetCI(d, "Tara", "TARA"));
                dto.ISO = AsStr(GetCI(d, "ISO", "iso"));
                dto.MAX_GROSS = ToDecimalSafe(GetCI(d, "maxgross", "MAXGROSS"));
                dto.LACRE_IPA = AsStr(GetCI(d, "Lacre_ipa", "LACRE_IPA"));
                dto.LACRE_IPA_TAG_ORIGEM = AsStr(GetCI(d, "lacre_origem", "LACRE_ORIGEM"));
                dto.LACRE1_GATE = AsStr(GetCI(d, "lacre1_gate", "LACRE1_GATE"));
                dto.LACRE1_TAG = AsStr(GetCI(d, "Lacre2", "LACRE2"));
                dto.LACRE2_GATE = AsStr(GetCI(d, "lacre2_gate", "LACRE2_GATE"));
                dto.LACRE2_TAG = AsStr(GetCI(d, "Lacre3", "LACRE3"));
                dto.LACRE3_GATE = AsStr(GetCI(d, "lacre3_gate", "LACRE3_GATE"));
                dto.LACRE3_TAG = AsStr(GetCI(d, "Lacre4", "LACRE4"));
                dto.LACRE_LONA = AsStr(GetCI(d, "lacre_lona", "LACRE_LONA"));
                dto.TEMP_SETPOINT = AsStr(GetCI(d, "TEMP", "temp"));
                dto.TEMP_MAX = AsStr(GetCI(d, "tempmax", "TEMPMAX"));
                dto.TEMP_MIN = AsStr(GetCI(d, "tempmin", "TEMPMIN"));

                var sb = new StringBuilder();
                if (dto.SISTEMA == "A")
                {
                    sb.AppendLine("CONTEINER TEMPORARIO");
                    var qArm = @"SELECT B.DESCR FROM SGIPA..TB_ARMAZENS_IPA A
                        INNER JOIN SGIPA..TB_CAD_TIPO_CT B ON A.TIPO_CT=B.AUTONUM WHERE A.AUTONUM=@A";
                    var ta = con.QueryFirstOrDefault<string>(qArm, new { A = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds());
                    sb.AppendLine(ta ?? "");
                    var qOc = @"SELECT REPLACE(ISNULL(A.PERC_OCUPACAO_CT,0),'%','') FROM SGIPA..TB_ARMAZENS_IPA A
                        INNER JOIN SGIPA..TB_CAD_TIPO_CT B ON A.TIPO_CT=B.AUTONUM WHERE A.AUTONUM=@A";
                    var oc = con.QueryFirstOrDefault<string>(qOc, new { A = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds());
                    sb.AppendLine("OCUPACAO: " + (oc ?? "0") + "%");
                }
                else if (dto.SISTEMA == "R")
                {
                    sb.AppendLine("CONTEINER EXPORTACAO");
                }
                else if (dto.SISTEMA == "I")
                {
                    sb.AppendLine("Imp:" + AsStr(GetCI(d, "Fantasia", "FANTASIA")));
                    var bloq = false;
                    if (dto.EF != "E")
                    {
                        if (ToInt32Safe(GetCI(d, "FLAG_BLOQUEIO_MANUAL", "flag_bloqueio_manual")) == 1) { bloq = true; sb.AppendLine("BLOQUEIO MANUAL DE CNTR"); }
                        if (ToInt32Safe(GetCI(d, "FLAG_BLOQUEIO", "flag_bloqueio")) == 1) { bloq = true; sb.AppendLine("BLOQUEIO ALFANDEGA (CNTR)"); }
                        if (ToInt32Safe(GetCI(d, "FLAG_ALERTA_COL_INDICADOR", "flag_alerta_col_indicador"), 1) == 1)
                        {
                            dto.ALERTA_LINHA_9 = "CARGA ESPECIAL : " + AsStr(GetCI(d, "Indicador", "INDICADOR"));
                            dto.ALERTA_LINHA_9_VISIVEL = true;
                        }

                        var qBlMan = @"SELECT COUNT(*) FROM sgipa..tb_bl bl
                            INNER JOIN sgipa..tb_amr_cntr_bl amr ON bl.autonum=amr.bl
                            WHERE amr.cntr=@C AND bl.flag_bloqueio_manual=1";
                        if (con.QuerySingle<int>(qBlMan, new { C = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds()) != 0)
                        { bloq = true; sb.AppendLine("BLOQUEIO MANUAL DE BL"); }

                        var qBlAlf = @"SELECT COUNT(*) FROM sgipa..tb_bl bl
                            INNER JOIN sgipa..tb_amr_cntr_bl amr ON bl.autonum=amr.bl
                            WHERE amr.cntr=@C AND bl.flag_bloqueio=1";
                        if (con.QuerySingle<int>(qBlAlf, new { C = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds()) != 0)
                        { bloq = true; sb.AppendLine("BLOQUEIO ALFANDEGA DE BL"); }

                        if (!bloq) sb.AppendLine("NAO CONSTA BLOQUEIO");

                        var qReg = @"SELECT COUNT(*) FROM sgipa..tb_registro_saida_cntr sc
                            INNER JOIN sgipa..tb_ordem_carregamento oc ON sc.ordem_carreg = oc.autonum
                            WHERE sc.cntr=@C AND ISNULL(oc.flag_saida,0)=0";
                        if (con.QuerySingle<int>(qReg, new { C = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds()) != 0)
                        {
                            sb.AppendLine("CARREGAMENTO REGISTRADO");
                            dto.ALERTA_LINHA_9 = "CARREGAMENTO REGISTRADO";
                            dto.ALERTA_LINHA_9_VISIVEL = true;
                        }
                        else sb.AppendLine("NAO CONSTA REGISTRO");

                        var ndoc = AsStr(GetCI(d, "Num_Documento", "NUM_DOCUMENTO"));
                        if (!string.IsNullOrWhiteSpace(ndoc))
                        {
                            sb.AppendLine("DOCUMENTO : " + ndoc);
                            var dtDes = GetCI(d, "Dt_Desembaraco", "DT_DESEMBARACO");
                            sb.AppendLine("Carga Desembaracada em " + FmtData(dtDes));
                        }

                        var sig = descrSig;
                        if (string.IsNullOrWhiteSpace(sig))
                        {
                            sig = con.QueryFirstOrDefault<string>(
                                "SELECT vw.descricao FROM sgipa.dbo.VW_SIGVIG_CNTR_ESTOQUE vw WHERE vw.CNTR = @C",
                                new { C = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds()) ?? "";
                        }
                        dto.SIGVIG = sig;

                        sb.AppendLine("Nao consta Agendamento");
                    }
                }

                dto.TEXTO_CONTEXTUAL = sb.ToString().TrimEnd();

                dto.LOTES_BL = con.Query<CntrPatioBlDto>(
                    @"SELECT tb_bl.autonum AS LOTE, tb_bl.numero AS BL
                      FROM sgipa..tb_bl
                      INNER JOIN sgipa..tb_amr_cntr_bl ON tb_bl.autonum = tb_amr_cntr_bl.bl
                      INNER JOIN sgipa..tb_cntr_bl ON tb_amr_cntr_bl.cntr = tb_cntr_bl.autonum
                      WHERE tb_cntr_bl.id_conteiner = @Id AND tb_bl.flag_ativo = 1",
                    new { Id = idConteiner },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();

                var stReefer = con.QueryFirstOrDefault(
                    @"SELECT ISNULL(a.flag_reefer_sem_temp,0) AS fts, ISNULL(a.flag_desligado,0) AS fd
                      FROM SGIPA..TB_CNTR_BL a WHERE a.AUTONUM=@A AND ISNULL(a.flag_reefer_sem_temp,0)<>0",
                    new { A = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds());
                if (stReefer != null)
                {
                    var fts = Convert.ToInt32(((IDictionary<string, object>)stReefer)["fts"]);
                    var fd = Convert.ToInt32(((IDictionary<string, object>)stReefer)["fd"]);
                    dto.AVISO_REEFER_POSICAO_BLOQUEADO = fts == 1 && fd == 0;
                    dto.AVISO_REEFER_POSICAO_MSG = dto.AVISO_REEFER_POSICAO_BLOQUEADO
                        ? "Conteiner Reefer sem temperatura - Posicionamento bloqueado"
                        : "Conteiner Reefer DESLIGADO";
                }

                var qTi = @"SELECT COUNT(1) FROM SGIPA..TB_CNTR_BL a
                    INNER JOIN SGIPA..TB_amr_CNTR_BL B ON a.autonum=b.cntr
                    INNER JOIN SGIPA..TB_BL c ON b.bl=c.autonum
                    INNER JOIN SGIPA..TB_cad_parceiros d ON c.importador=d.autonum
                    WHERE a.AUTONUM=@A AND d.flag_transp_interno=1";
                dto.FRAME_TRANSPORTE_INTERNO = con.QuerySingle<int>(qTi, new { A = dto.AUTONUM }, commandTimeout: Config.QueryTimeoutInSeconds()) != 0;
            }

            return dto;
        }

        public List<ProximoMovimentoDto> ProximosMovimentos(string idConteiner)
        {
            idConteiner = (idConteiner ?? string.Empty).Trim().ToUpperInvariant();
            using (var con = Open())
            {
                var sql = @"
SELECT MOTIVO, LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV
FROM OPERADOR..VW_PROX_MOV_CNTR WHERE ID_CONTEINER=@Id
UNION
SELECT MOTIVO, ' ' AS DT_PREV FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE ID_CONTEINER=@Id";
                return con.Query<ProximoMovimentoDto>(sql, new { Id = idConteiner }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<VizinhoPilhaDto> VizinhancaPilha(VizinhancaPilhaRequest request, int patioUsuario)
        {
            request = request ?? new VizinhancaPilhaRequest();
            var id = (request.ID_CONTEINER ?? "").Trim().ToUpperInvariant();
            var posBusca = request.MODO_DESTINO
                ? (request.PREFIXO_DESTINO ?? "").Trim().ToUpperInvariant()
                : (request.YARD_ATUAL ?? "").Trim().ToUpperInvariant();

            if (string.IsNullOrEmpty(posBusca)) return new List<VizinhoPilhaDto>();

            using (var con = Open())
            {
                TbYardVal yardRow = null;
                if (patioUsuario != 3)
                {
                    yardRow = con.QueryFirstOrDefault<TbYardVal>(
                        "SELECT yard AS Yard, valida AS Valida FROM operador..tb_yard WHERE patio IN (1,7) AND yard=@Y",
                        new { Y = posBusca }, commandTimeout: Config.QueryTimeoutInSeconds());
                }
                else
                {
                    yardRow = con.QueryFirstOrDefault<TbYardVal>(
                        @"SELECT TOP 1 yard AS Yard, valida AS Valida FROM operador..tb_yard
                          WHERE patio=3 AND (yard=@Y OR yard_eq=@Y)",
                        new { Y = posBusca }, commandTimeout: Config.QueryTimeoutInSeconds());
                }

                var patioSql = PatioSqlFragment(patioUsuario);
                if (yardRow != null && !string.IsNullOrEmpty(yardRow.Yard))
                {
                    posBusca = yardRow.Yard;
                    if (yardRow.Valida == 1)
                    {
                        var sql = BuildGrid2ValidaSql(patioUsuario, request.MODO_DESTINO, posBusca, id);
                        return con.Query<VizinhoPilhaDto>(sql.Sql, sql.Param, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
                    }

                    var sqlEq = @"
SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV,MOTIVO
FROM sgipa..VW_PROX_MOV_CNTR_ALL WHERE YARD=@Y AND id_conteiner<>@Id
UNION
SELECT ID_CONTEINER,YARD,NULL AS DT_PREV,MOTIVO FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE YARD=@Y AND id_conteiner<>@Id
ORDER BY ID_CONTEINER";
                    return con.Query<VizinhoPilhaDto>(sqlEq, new { Y = posBusca, Id = id }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
                }

                var sqlAmb = BuildGrid2AmbiguoSql(patioUsuario, request.MODO_DESTINO, posBusca, id);
                return con.Query<VizinhoPilhaDto>(sqlAmb.Sql, sqlAmb.Param, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        private sealed class TbYardVal { public string Yard { get; set; } public int Valida { get; set; } }

        private static string PatioSqlFragment(int patioUsuario)
        {
            return patioUsuario == 1 || patioUsuario == 7 ? " PATIO IN (1,7) " : " PATIO = " + patioUsuario + " ";
        }

        private static SqlBuild BuildGrid2ValidaSql(int patioUsuario, bool modoDestino, string posBusca, string idConteiner)
        {
            var p = new DynamicParameters();
            p.Add("Id", idConteinarFix(idConteiner));
            var patioFrag = PatioSqlFragment(patioUsuario);
            string likePat;
            if (!modoDestino)
            {
                likePat = patioUsuario != 3
                    ? (posBusca.Length >= 3 ? posBusca.Substring(0, 3) : posBusca) + "%"
                    : (posBusca.Length >= 4 ? posBusca.Substring(0, 4) : posBusca) + "%";
            }
            else
            {
                if (posBusca.Length == 1) likePat = posBusca.Substring(0, 1) + "%";
                else if (posBusca.Length == 2) likePat = posBusca.Substring(0, 2) + "%";
                else
                    likePat = patioUsuario != 3
                        ? (posBusca.Length >= 3 ? posBusca.Substring(0, 3) : posBusca) + "%"
                        : (posBusca.Length >= 4 ? posBusca.Substring(0, 4) : posBusca) + "%";
            }

            p.Add("LikePat", likePat);

            string subPilha = patioUsuario != 3 ? "SUBSTRING(YARD,2,2)" : "SUBSTRING(YARD,3,2)";
            var sb = new StringBuilder();
            sb.Append("SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV,MOTIVO,");
            sb.Append(subPilha).Append(" AS PILHA FROM sgipa..VW_PROX_MOV_CNTR_all WHERE ").Append(patioFrag).Append(" AND YARD LIKE @LikePat AND id_conteiner<>@Id");
            sb.Append(" UNION ");
            sb.Append("SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV,MOTIVO,");
            sb.Append(subPilha).Append(" AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE ").Append(patioFrag).Append(" AND YARD LIKE @LikePat AND id_conteiner<>@Id");
            sb.Append(" ORDER BY PILHA , ID_CONTEINER ");
            return new SqlBuild { Sql = sb.ToString(), Param = p };
        }

        private static string idConteinarFix(string id) => (id ?? "").Trim().ToUpperInvariant();

        private static SqlBuild BuildGrid2AmbiguoSql(int patioUsuario, bool modoDestino, string posBusca, string idConteiner)
        {
            var p = new DynamicParameters();
            p.Add("Id", idConteinarFix(idConteiner));
            var patioFrag = PatioSqlFragment(patioUsuario);
            string subPilha = patioUsuario != 3 ? "SUBSTRING(YARD,2,2)" : "SUBSTRING(YARD,3,2)";

            var yardClause = new StringBuilder();
            if (posBusca.Length == 1)
                yardClause.Append(" AND YARD LIKE @P1 ");
            else if (posBusca.Length == 2)
                yardClause.Append(" AND YARD LIKE @P2 ");
            else if (posBusca.Length >= 3)
            {
                if (patioUsuario != 3)
                {
                    var mid = posBusca.Length >= 4 ? posBusca.Substring(1, 2) : "";
                    int v;
                    if (mid.Length == 2 && int.TryParse(mid, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v != 0)
                    {
                        if (v % 2 != 0)
                        {
                            yardClause.Append(" AND (YARD LIKE @P3a OR YARD LIKE @P3b OR YARD LIKE @P3c) ");
                            p.Add("P3a", posBusca.Substring(0, 3) + "%");
                            p.Add("P3b", posBusca.Substring(0, 1) + (v + 1).ToString("00", CultureInfo.InvariantCulture) + "%");
                            p.Add("P3c", posBusca.Substring(0, 1) + (v - 1).ToString("00", CultureInfo.InvariantCulture) + "%");
                        }
                        else
                        {
                            yardClause.Append(" AND (YARD LIKE @P4a OR YARD LIKE @P4b OR YARD LIKE @P4c OR YARD LIKE @P4d) ");
                            p.Add("P4a", posBusca.Substring(0, 3) + "%");
                            p.Add("P4b", posBusca.Substring(0, 1) + (v + 1).ToString("00", CultureInfo.InvariantCulture) + "%");
                            p.Add("P4c", posBusca.Substring(0, 1) + (v + 2).ToString("00", CultureInfo.InvariantCulture) + "%");
                            p.Add("P4d", posBusca.Substring(0, 1) + (v - 2).ToString("00", CultureInfo.InvariantCulture) + "%");
                        }
                    }
                    else
                    {
                        yardClause.Append(" AND YARD LIKE @P5 ");
                        p.Add("P5", posBusca.Substring(0, 3) + "%");
                    }
                }
                else
                {
                    var mid = posBusca.Length >= 5 ? posBusca.Substring(2, 2) : "";
                    int v;
                    if (mid.Length == 2 && int.TryParse(mid, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v != 0)
                    {
                        if (v % 2 != 0)
                        {
                            yardClause.Append(" AND (YARD LIKE @C3a OR YARD LIKE @C3b OR YARD LIKE @C3c) ");
                            p.Add("C3a", posBusca.Substring(0, 4) + "%");
                            p.Add("C3b", posBusca.Substring(0, 2) + (v + 1).ToString("00", CultureInfo.InvariantCulture) + "%");
                            p.Add("C3c", posBusca.Substring(0, 2) + (v - 1).ToString("00", CultureInfo.InvariantCulture) + "%");
                        }
                        else
                        {
                            yardClause.Append(" AND (YARD LIKE @C4a OR YARD LIKE @C4b OR YARD LIKE @C4c OR YARD LIKE @C4d) ");
                            p.Add("C4a", posBusca.Substring(0, 4) + "%");
                            p.Add("C4b", posBusca.Substring(0, 2) + (v + 1).ToString("00", CultureInfo.InvariantCulture) + "%");
                            p.Add("C4c", posBusca.Substring(0, 2) + (v + 2).ToString("00", CultureInfo.InvariantCulture) + "%");
                            p.Add("C4d", posBusca.Substring(0, 2) + (v - 2).ToString("00", CultureInfo.InvariantCulture) + "%");
                        }
                    }
                    else
                    {
                        yardClause.Append(" AND YARD LIKE @C5 ");
                        p.Add("C5", posBusca.Substring(0, 4) + "%");
                    }
                }
            }

            if (posBusca.Length == 1) p.Add("P1", posBusca.Substring(0, 1) + "%");
            if (posBusca.Length == 2) p.Add("P2", posBusca.Substring(0, 2) + "%");

            var sb = new StringBuilder();
            sb.Append("SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV,MOTIVO,");
            sb.Append(subPilha).Append(" AS PILHA FROM SGIPA..VW_PROX_MOV_CNTR_ALL WHERE ").Append(patioFrag).Append(yardClause).Append(" AND id_conteiner<>@Id ");
            sb.Append(" UNION ");
            sb.Append("SELECT ID_CONTEINER,YARD,NULL AS DT_PREV,MOTIVO,").Append(subPilha).Append(" AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE ").Append(patioFrag).Append(yardClause).Append(" AND id_conteiner<>@Id ");
            sb.Append(" ORDER BY PILHA , ID_CONTEINER ");

            return new SqlBuild { Sql = sb.ToString(), Param = p };
        }

        private sealed class SqlBuild { public string Sql { get; set; } public DynamicParameters Param { get; set; } }

        #endregion

        #region Movimentação

        public MovimentacaoCntrConfirmResult ConfirmarMovimentacao(MovimentacaoCntrConfirmRequest req, int patioUsuario, int usuarioId)
        {
            req = req ?? new MovimentacaoCntrConfirmRequest();
            var res = new MovimentacaoCntrConfirmResult();

            if (req.AUTONUM_CNTR <= 0)
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Um conteiner deve ser informado para pesquisa!";
                return res;
            }

            var destino = (req.YARD_DESTINO ?? "").Trim().ToUpperInvariant();
            var origem = (req.YARD_ORIGEM ?? "").Trim().ToUpperInvariant();
            var sistema = (req.SISTEMA ?? "").Trim().ToUpperInvariant();
            var idCntr = (req.ID_CONTEINER ?? "").Trim().ToUpperInvariant();

            using (var con = Open())
            {
                var reeferBloq = con.QueryFirstOrDefault(
                    @"SELECT ISNULL(flag_reefer_sem_temp,0) AS fts, ISNULL(flag_desligado,0) AS fd
                      FROM SGIPA.DBO.TB_CNTR_BL WHERE AUTONUM=@A AND ISNULL(flag_reefer_sem_temp,0)<>0",
                    new { A = req.AUTONUM_CNTR }, commandTimeout: Config.QueryTimeoutInSeconds());
                if (reeferBloq != null)
                {
                    var dict = (IDictionary<string, object>)reeferBloq;
                    var fts = Convert.ToInt32(dict["fts"]);
                    var fd = Convert.ToInt32(dict["fd"]);
                    if (fts == 1 && fd == 0)
                    {
                        res.SUCESSO = false;
                        res.MENSAGEM = "Conteiner Reefer sem temperatura - Posicionamento bloqueado";
                        return res;
                    }
                }

                if (sistema == "I")
                {
                    var cntHist = con.QuerySingle<int>(
                        @"SELECT COUNT(1) FROM OPERADOR..TB_HIST_SHIFTING WHERE CNTR=@C AND TIPO='I' AND ORIGEM<>'GATEIN'",
                        new { C = req.AUTONUM_CNTR }, commandTimeout: Config.QueryTimeoutInSeconds());
                    if (cntHist == 0 && !req.LACRES_CONFERIDOS)
                    {
                        res.SUCESSO = false;
                        res.MENSAGEM = "Informe que os lacres foram conferidos";
                        return res;
                    }
                }

                if (sistema == "I")
                {
                    var lacresOk = !string.IsNullOrWhiteSpace(req.LACRE_IPA)
                        || !string.IsNullOrWhiteSpace(req.LACRE1_GATE)
                        || !string.IsNullOrWhiteSpace(req.LACRE2_GATE)
                        || !string.IsNullOrWhiteSpace(req.LACRE3_GATE);
                    if (!lacresOk)
                    {
                        res.SUCESSO = false;
                        res.MENSAGEM = "Pelo menos 1 lacre de origem deve ser informado";
                        return res;
                    }
                }

                if (sistema == "I" && req.EF == "F"
                    && string.IsNullOrWhiteSpace(req.LACRE_LONA)
                    && string.Equals(req.TIPO_BASICO, "OT", StringComparison.OrdinalIgnoreCase))
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = "Informe o Lacre Lona ";
                    return res;
                }

                if (req.MOTIVO_ID <= 0)
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = "Informe o motivo da movimentacao!";
                    return res;
                }

                if (MotivoExigeCamera(req.MOTIVO_ID) && (!req.CAMERA_ID.HasValue || req.CAMERA_ID.Value <= 0))
                {
                    res.SUCESSO = false;
                    res.EXIGE_CAMERA = true;
                    res.MENSAGEM = "Selecione uma camera";
                    return res;
                }

                if (string.IsNullOrEmpty(destino))
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = "Destino nao Informado!";
                    return res;
                }

                if (string.Equals(destino, origem, StringComparison.OrdinalIgnoreCase))
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = "Destino e Origem iguais!";
                    return res;
                }

                if (destino == "SAIDA" && !req.CONFIRMOU_SAIDA)
                {
                    res.SUCESSO = false;
                    res.EXIGE_CONFIRMACAO_SAIDA = true;
                    res.MENSAGEM = "Confirme a saida da unidade";
                    return res;
                }

                if (destino == "SAIDA" && sistema == "I")
                {
                    var bloqObj = con.ExecuteScalar<object>(
                        "SELECT SGIPA.DBO.FN_BLOQUEIO_SCANNER(@C,'SAIDA')",
                        new { C = req.AUTONUM_CNTR }, commandTimeout: Config.QueryTimeoutInSeconds());
                    var bloqScan = bloqObj != null && Convert.ToBoolean(bloqObj, CultureInfo.InvariantCulture);
                    if (bloqScan)
                    {
                        res.SUCESSO = false;
                        res.MENSAGEM = "Esta unidade tem BLOQUEIO de ESCANER";
                        return res;
                    }
                }

                var bkMsg = ValidarBk(con, req.AUTONUM_CNTR, sistema, destino);
                if (bkMsg != null) { res.SUCESSO = false; res.MENSAGEM = bkMsg; return res; }

                var cliaMsg = ValidarClia(con, req.AUTONUM_CNTR, sistema, destino);
                if (cliaMsg != null) { res.SUCESSO = false; res.MENSAGEM = cliaMsg; return res; }

                if (!string.Equals(destino, "DTA", StringComparison.OrdinalIgnoreCase))
                {
                    var lacMsg = ValidarSequenciaLacres(req);
                    if (lacMsg != null) { res.SUCESSO = false; res.MENSAGEM = lacMsg; return res; }
                }

                if (req.EF == "E" && destino == "SAIDA" && !req.CONFIRMOU_VEICULO_SERV)
                {
                    res.SUCESSO = false;
                    res.EXIGE_VEICULO_SERV = true;
                    res.MENSAGEM = "Confirme o fluxo de veiculo/servico (equivalente ao modal legado)";
                    return res;
                }

                if (destino == "DTA" && !req.CONFIRMOU_VEICULO_SERV)
                {
                    res.SUCESSO = false;
                    res.EXIGE_VEICULO_SERV = true;
                    res.MENSAGEM = "Confirme o fluxo DTA / veiculo-servico antes de movimentar";
                    return res;
                }

                if (destino == "DTA")
                    destino = "CAM";

                var fnSql = @"SELECT [Operador].[dbo].[fnMovimentaCntr](@Sigla,@ColPatio,@Tam,@Sistema,@YardO,@YardD,@Usuario,@FlagBloq)";
                var fnMsgObj = con.ExecuteScalar<object>(fnSql, new
                {
                    Sigla = idCntr,
                    ColPatio = patioUsuario,
                    Tam = req.TAMANHO,
                    Sistema = sistema,
                    YardO = origem,
                    YardD = destino,
                    Usuario = usuarioId,
                    FlagBloq = req.FLAG_CNTR_BLOQ_CHAO_PATIO
                }, commandTimeout: Config.QueryTimeoutInSeconds());
                var fnMsg = fnMsgObj == null ? "" : Convert.ToString(fnMsgObj, CultureInfo.InvariantCulture).Trim();

                if (string.IsNullOrEmpty(fnMsg) || !fnMsg.Equals("OK", StringComparison.OrdinalIgnoreCase))
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = string.IsNullOrEmpty(fnMsg) ? "Validacao de movimentacao falhou" : fnMsg;
                    return res;
                }

                ExecutarPcMovimentaCntr(con, req, usuarioId, destino, sistema, idCntr, origem);

                ProcessarDivergenciaLacres(con, req);

                if (!string.IsNullOrWhiteSpace(req.IMO))
                {
                    var segreg = VerificaRegrasSegImoDelta(con, idCntr, destino, patioUsuario);
                    if (!string.IsNullOrWhiteSpace(segreg))
                    {
                        ExecutarPcMovimentaCntrEstorno(con, req, usuarioId, destino, sistema, idCntr, origem);
                        res.SUCESSO = false;
                        res.MENSAGEM = "SEGREGACAO IMO !!!! " + segreg;
                        return res;
                    }
                }

                res.SUCESSO = true;
                res.MENSAGEM = "Movimentacao concluida";
                return res;
            }
        }

        private static string ValidarSequenciaLacres(MovimentacaoCntrConfirmRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.LACRE_IPA))
            {
                if (!string.IsNullOrWhiteSpace(req.LACRE1_GATE) || !string.IsNullOrWhiteSpace(req.LACRE2_GATE) || !string.IsNullOrWhiteSpace(req.LACRE3_GATE))
                    return "Lacres informados incorretamente";
            }
            if (string.IsNullOrWhiteSpace(req.LACRE1_GATE))
            {
                if (!string.IsNullOrWhiteSpace(req.LACRE2_GATE) || !string.IsNullOrWhiteSpace(req.LACRE3_GATE))
                    return "Lacres informados incorretamente";
            }
            if (string.IsNullOrWhiteSpace(req.LACRE2_GATE))
            {
                if (!string.IsNullOrWhiteSpace(req.LACRE3_GATE))
                    return "Lacres informados incorretamente";
            }
            return null;
        }

        private static string ValidarBk(SqlConnection con, long autonum, string sistema, string destino)
        {
            if (!string.Equals(destino, "BK", StringComparison.OrdinalIgnoreCase) || autonum <= 0 || sistema != "I")
                return null;

            var flag = con.ExecuteScalar<int?>(
                "SELECT ISNULL(FLAG_SEL_TRANSF_BK,0) FROM SGIPA..TB_CNTR_BL WHERE AUTONUM=@A",
                new { A = autonum }, commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
            if (flag == 0)
                return "Unidade sem autorizacao para Transferencia BK";

            var sqlAuto = @"
SELECT ISNULL(MAX(C.AUTONUM_GD_RESERVA),0) FROM SGIPA..TB_CNTR_BL A
INNER JOIN OPERADOR..TB_GD_RESERVA B ON A.AUTONUM_GD_RESERVA=B.AUTONUM_GD_RESERVA
INNER JOIN OPERADOR..TB_GD_RESERVA C ON B.PERIODO_INICIAL=C.PERIODO_INICIAL AND B.SERVICO_GATE=C.SERVICO_GATE AND C.PATIO=2
LEFT JOIN (
    SELECT COUNT(1) AG, AUTONUM_GD_RESERVA FROM SGIPA..TB_CNTR_BL WHERE FLAG_TERMINAL=1 AND FLAG_HISTORICO=0 GROUP BY AUTONUM_GD_RESERVA
) D ON C.AUTONUM_GD_RESERVA=D.AUTONUM_GD_RESERVA OR B.PATIO=2
WHERE ISNULL(C.LIMITE_MOVIMENTOS,0)>ISNULL(D.AG,0) AND AUTONUM=@A";
            var auto = con.ExecuteScalar<long?>(sqlAuto, new { A = autonum }, commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
            if (auto > 0)
            {
                con.Execute(
                    "UPDATE SGIPA..TB_CNTR_BL SET AUTONUM_GD_RESERVA=@R WHERE AUTONUM=@A",
                    new { R = auto, A = autonum }, commandTimeout: Config.QueryTimeoutInSeconds());
                return null;
            }

            var cntAg = con.ExecuteScalar<int?>(
                "SELECT COUNT(1) FROM SGIPA..VW_AGG_AGENDAMENTOS WHERE AUTONUMCNTR=@A",
                new { A = autonum }, commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
            if (cntAg != 0)
            {
                var texto = con.ExecuteScalar<string>(
                    @"SELECT ISNULL(MAX('Limite Janela BK ' + CONVERT(VARCHAR,ISNULL(C.LIMITE_MOVIMENTOS,0))+' Agendados ' + CONVERT(VARCHAR,ISNULL(D.AG,0)) ),'')
                      FROM SGIPA..TB_CNTR_BL A
                      INNER JOIN OPERADOR..TB_GD_RESERVA B ON A.AUTONUM_GD_RESERVA=B.AUTONUM_GD_RESERVA
                      INNER JOIN OPERADOR..TB_GD_RESERVA C ON B.PERIODO_INICIAL=C.PERIODO_INICIAL AND B.SERVICO_GATE=C.SERVICO_GATE AND C.PATIO=2
                      LEFT JOIN (
                          SELECT COUNT(1) AG, AUTONUM_GD_RESERVA FROM SGIPA..TB_CNTR_BL WHERE FLAG_TERMINAL=1 AND FLAG_HISTORICO=0 GROUP BY AUTONUM_GD_RESERVA
                      ) D ON C.AUTONUM_GD_RESERVA=D.AUTONUM_GD_RESERVA
                      WHERE AUTONUM=@A",
                    new { A = autonum }, commandTimeout: Config.QueryTimeoutInSeconds());
                return "Unidade com Agendamento, Transferencia BK nao autorizada " + (texto ?? "");
            }

            return null;
        }

        private static string ValidarClia(SqlConnection con, long autonum, string sistema, string destino)
        {
            if (!string.Equals(destino, "CLIA", StringComparison.OrdinalIgnoreCase) || autonum <= 0 || sistema != "I")
                return null;
            var flag = con.ExecuteScalar<int?>(
                "SELECT ISNULL(FLAG_SEL_TRANSF_CLIA,0) FROM SGIPA..TB_CNTR_BL WHERE AUTONUM=@A",
                new { A = autonum }, commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
            if (flag == 0)
                return "Unidade sem autorizacao para Transferencia CLIA";
            return null;
        }

        private static void ExecutarPcMovimentaCntr(SqlConnection con, MovimentacaoCntrConfirmRequest req, int usuarioId, string yardDestino, string sistema, string idCntr, string yardOrigem)
        {
            var p = new DynamicParameters();
            p.Add("Autonum", req.AUTONUM_CNTR);
            p.Add("Sigla", idCntr);
            p.Add("Gmci", (req.GMCI ?? "").Trim());
            p.Add("Sistema", sistema);
            p.Add("YardOrigem", yardOrigem);
            p.Add("YardDestino", yardDestino);
            p.Add("Tara", (int)req.TARA);
            p.Add("Lacre_Ipa", (req.LACRE_IPA ?? "").Trim());
            p.Add("Lacre1_Gate", (req.LACRE1_GATE ?? "").Trim());
            p.Add("Lacre2_Gate", (req.LACRE2_GATE ?? "").Trim());
            p.Add("Lacre3_Gate", (req.LACRE3_GATE ?? "").Trim());
            p.Add("Lacre_Lona", (req.LACRE_LONA ?? "").Trim());
            p.Add("Camera_Posicionamento", req.CAMERA_ID ?? 0);
            p.Add("Usuario", usuarioId);
            p.Add("Motivo", req.MOTIVO_ID);

            con.Execute(
                "operador.dbo.pcMovimentaCntr",
                p,
                commandTimeout: Config.QueryTimeoutInSeconds(),
                commandType: CommandType.StoredProcedure);
        }

        private static void ExecutarPcMovimentaCntrEstorno(SqlConnection con, MovimentacaoCntrConfirmRequest req, int usuarioId, string yardDestino, string sistema, string idCntr, string yardOrigem)
        {
            var p = new DynamicParameters();
            p.Add("Autonum", req.AUTONUM_CNTR);
            p.Add("Sigla", idCntr);
            p.Add("Gmci", (req.GMCI ?? "").Trim());
            p.Add("Sistema", sistema);
            p.Add("YardOrigem", yardDestino);
            p.Add("YardDestino", yardOrigem);
            p.Add("Tara", (int)req.TARA);
            p.Add("Lacre_Ipa", (req.LACRE_IPA ?? "").Trim());
            p.Add("Lacre1_Gate", (req.LACRE1_GATE ?? "").Trim());
            p.Add("Lacre2_Gate", (req.LACRE2_GATE ?? "").Trim());
            p.Add("Lacre3_Gate", (req.LACRE3_GATE ?? "").Trim());
            p.Add("Lacre_Lona", (req.LACRE_LONA ?? "").Trim());
            p.Add("Camera_Posicionamento", req.CAMERA_ID ?? 0);
            p.Add("Usuario", usuarioId);
            p.Add("Motivo", req.MOTIVO_ID);

            con.Execute(
                "operador.dbo.pcMovimentaCntr",
                p,
                commandTimeout: Config.QueryTimeoutInSeconds(),
                commandType: CommandType.StoredProcedure);
        }

        private static void ProcessarDivergenciaLacres(SqlConnection con, MovimentacaoCntrConfirmRequest req)
        {
            if (!string.Equals(req.SISTEMA, "I", StringComparison.OrdinalIgnoreCase)) return;

            var lacresManif = "***" + (req.LACRE_IPA_TAG_ORIGEM ?? "").Trim()
                + "***" + (req.LACRE1_TAG ?? "").Trim()
                + "***" + (req.LACRE2_TAG ?? "").Trim()
                + "***" + (req.LACRE3_TAG ?? "").Trim();
            lacresManif = lacresManif.ToUpperInvariant();

            bool div = true;
            var ipa = (req.LACRE_IPA ?? "").Trim();
            var l1 = (req.LACRE1_GATE ?? "").Trim();
            var l2 = (req.LACRE2_GATE ?? "").Trim();
            var l3 = (req.LACRE3_GATE ?? "").Trim();

            if (!string.IsNullOrEmpty(ipa) && lacresManif.IndexOf(ipa.ToUpperInvariant(), StringComparison.Ordinal) >= 0) div = false;
            if (!string.IsNullOrEmpty(l1) && lacresManif.IndexOf(l1.ToUpperInvariant(), StringComparison.Ordinal) >= 0) div = false;
            if (!string.IsNullOrEmpty(l2) && lacresManif.IndexOf(l2.ToUpperInvariant(), StringComparison.Ordinal) >= 0) div = false;
            if (!string.IsNullOrEmpty(l3) && lacresManif.IndexOf(l3.ToUpperInvariant(), StringComparison.Ordinal) >= 0) div = false;

            if (!div) return;

            const string Xlocal = "12";
            const string Xtipo = "32";
            const string Xcompl = "8";
            var complIpa = LeftStr(ipa + "|" + l1 + "|" + l2 + "|" + l3, 35);

            var existe = con.ExecuteScalar<int?>(
                @"SELECT autonum FROM sgipa..tb_avarias_conteiner WHERE local=@L AND tipo=@T AND complemento=@C AND cntr=@Cnt",
                new { L = Xlocal, T = Xtipo, C = Xcompl, Cnt = req.AUTONUM_CNTR }, commandTimeout: Config.QueryTimeoutInSeconds());
            if (existe != null && existe > 0) return;

            con.Execute(
                @"INSERT INTO sgipa..TB_AVARIAS_CONTEINER (CNTR,LOCAL,TIPO,COMPLEMENTO,COMPL_IPA,PROVIDENCIA)
                  VALUES (@Cnt,@L,@T,@C,@Compl,11)",
                new { Cnt = req.AUTONUM_CNTR, L = Xlocal, T = Xtipo, C = Xcompl, Compl = complIpa },
                commandTimeout: Config.QueryTimeoutInSeconds());
        }

        private static string LeftStr(string s, int n)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Length <= n ? s : s.Substring(0, n);
        }

        private string VerificaRegrasSegImoDelta(SqlConnection con, string idConteiner, string yard, int patioUsuario)
        {
            var pUse = patioUsuario == 7 ? 1 : patioUsuario;
            var ep = con.QueryFirstOrDefault(
                "SELECT DIST_IMO_SEGREG1, DIST_IMO_SEGREG2, DIST_IMO_SEGREG3, DIST_IMO_SEGREG4 FROM sgipa..dte_tb_parametros",
                commandTimeout: Config.QueryTimeoutInSeconds());
            if (ep == null) return "";
            var d = (IDictionary<string, object>)ep;
            var e1 = ToDecimalSafe(GetCI(d, "DIST_IMO_SEGREG1", "dist_imo_segreg1"));
            var e2 = ToDecimalSafe(GetCI(d, "DIST_IMO_SEGREG2", "dist_imo_segreg2"));
            var e3 = ToDecimalSafe(GetCI(d, "DIST_IMO_SEGREG3", "dist_imo_segreg3"));
            var e4 = ToDecimalSafe(GetCI(d, "DIST_IMO_SEGREG4", "dist_imo_segreg4"));

            var vw = pUse == 3 ? "OPERADOR..VW_DIST_IMO_CLIA" : "OPERADOR..VW_DIST_IMO";
            var sqlRows = $"SELECT ID_CONTEINER,IMO1,IMO2,IMO3,IMO4,ID_CONTEINER_D,YARD_D,IMO1_D,IMO2_D,IMO3_D,IMO4_D,DIST_DELTA FROM {vw} WHERE ID_CONTEINER=@Id ORDER BY DIST_DELTA";
            var rows = con.Query(sqlRows, new { Id = idConteiner }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            var sb = new StringBuilder();

            foreach (var row in rows)
            {
                var rowDict = (IDictionary<string, object>)row;
                var distDelta = ToDecimalSafe(GetCI(rowDict, "DIST_DELTA", "dist_delta"));

                for (var i = 1; i <= 4; i++)
                {
                    for (var j = 1; j <= 4; j++)
                    {
                        var imoi = NormImo(AsStr(GetCI(rowDict, "IMO" + i, "imo" + i)));
                        var imoj = NormImo(AsStr(GetCI(rowDict, "IMO" + j + "_D", "imo" + j + "_d")));
                        if (string.IsNullOrEmpty(imoi) || string.IsNullOrEmpty(imoj) || imoi == "0" || imoj == "0") continue;

                        var seg = con.QueryFirstOrDefault<string>(
                            @"SELECT TOP 1 SEGREGACAO FROM (
                                SELECT SEGREGACAO FROM OPERADOR..tb_segregacao WHERE CLASS1=@A AND CLASS2=@B AND SEGREGACAO IN ('1','2','3','4')
                                UNION
                                SELECT SEGREGACAO FROM OPERADOR..tb_segregacao WHERE CLASS2=@A AND CLASS1=@B AND SEGREGACAO IN ('1','2','3','4')
                              ) x",
                            new { A = imoi, B = imoj }, commandTimeout: Config.QueryTimeoutInSeconds());

                        if (string.IsNullOrEmpty(seg)) continue;
                        var idx = int.Parse(seg, CultureInfo.InvariantCulture);
                        var min = idx == 1 ? e1 : idx == 2 ? e2 : idx == 3 ? e3 : e4;
                        if (distDelta < min)
                        {
                            var idD = AsStr(GetCI(rowDict, "ID_CONTEINER_D", "id_conteiner_d"));
                            var yardD = AsStr(GetCI(rowDict, "YARD_D", "yard_d"));
                            sb.Append("Espaco inferior a ").Append(min.ToString(CultureInfo.InvariantCulture)).Append("m ")
                                .Append(" Conteiner: ").Append(idD).Append(" Posicao: ").Append(yardD)
                                .Append(" Imo:").Append(imoj).AppendLine();
                        }
                    }
                }
            }

            return sb.ToString().TrimEnd();
        }

        private static string NormImo(string s)
        {
            return (s ?? "").Replace(".", "").Trim();
        }

        #endregion

        #region Histórico / Avarias

        public List<HistoricoShiftingRowDto> HistoricoShifting(string idConteiner, string sistema)
        {
            idConteiner = (idConteiner ?? "").Trim().ToUpperInvariant();
            sistema = (sistema ?? "").Trim().ToUpperInvariant();
            string sql;
            if (sistema == "R")
            {
                sql = @"SELECT CONVERT(VARCHAR(19), hs.data, 120) AS DATA, hs.origem AS ORIGEM, hs.destino AS DESTINO,
                        usu.usuario AS NMUSUARIO, cm.descricao AS DESC_MOTIVO, '' AS IDENTIFICACAO
                        FROM redex..tb_patio cc
                        INNER JOIN operador..tb_hist_shifting hs ON cc.autonum_patio = hs.cntr AND hs.tipo='R'
                        LEFT JOIN redex..tb_cad_usuarios usu ON hs.usuario = usu.autonum_usu
                        LEFT JOIN operador..tb_cad_motivo cm ON hs.motivo = cm.autonum
                        WHERE cc.id_conteiner=@Id ORDER BY hs.data DESC";
            }
            else if (sistema == "A")
            {
                sql = @"SELECT CONVERT(VARCHAR(19), hs.data, 120) AS DATA, hs.origem AS ORIGEM, hs.destino AS DESTINO,
                        usu.usuario AS NMUSUARIO, cm.descricao AS DESC_MOTIVO, '' AS IDENTIFICACAO
                        FROM sgipa..tb_armazens_ipa cc
                        INNER JOIN operador..tb_hist_shifting hs ON cc.autonum = hs.cntr AND hs.tipo='A'
                        LEFT JOIN sgipa..tb_cad_usuarios usu ON hs.usuario = usu.autonum
                        LEFT JOIN operador..tb_cad_motivo cm ON hs.motivo = cm.autonum
                        WHERE cc.descr=@Id ORDER BY hs.data DESC";
            }
            else
            {
                sql = @"SELECT CONVERT(VARCHAR(19), hs.data, 120) AS DATA, hs.origem AS ORIGEM, hs.destino AS DESTINO,
                        usu.usuario AS NMUSUARIO, cm.descricao AS DESC_MOTIVO, '' AS IDENTIFICACAO
                        FROM sgipa..tb_cntr_bl cc
                        INNER JOIN operador..tb_hist_shifting hs ON cc.autonum = hs.cntr AND hs.tipo='I'
                        LEFT JOIN sgipa..tb_cad_usuarios usu ON hs.usuario = usu.autonum
                        LEFT JOIN operador..tb_cad_motivo cm ON hs.motivo = cm.autonum
                        WHERE cc.id_conteiner=@Id ORDER BY hs.data DESC";
            }

            using (var con = Open())
            {
                return con.Query<HistoricoShiftingRowDto>(sql, new { Id = idConteiner }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<AvariaHistDto> HistoricoAvarias(long autonumCntr)
        {
            using (var con = Open())
            {
                var sql = @"
SELECT RTRIM(ISNULL(CONVERT(CHAR, A.DT_CAD_AVARIA,103),'')) + ' ' + ISNULL(CONVERT(CHAR, A.DT_CAD_AVARIA,108),'') AS DT_AVARIA,
       (RTRIM(B.DESCR) + ' ' + RTRIM(C.DESCR)) AS AVARIA,
       ISNULL(u.usuario,' ') AS USUARIO
FROM sgipa.dbo.tb_avarias_conteiner a
INNER JOIN sgipa.dbo.dte_tb_avarias b ON a.local=b.code AND B.ident='L'
INNER JOIN sgipa.dbo.dte_tb_avarias c ON a.tipo=c.code AND C.IDENT='T'
LEFT JOIN sgipa.dbo.tb_cad_usuarios u ON a.usu_cad_avaria=u.autonum
WHERE a.CNTR = @C";
                return con.Query<AvariaHistDto>(sql, new { C = autonumCntr }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<AvariaCntrGridDto> ListarAvariasAbertas(long autonumCntr)
        {
            using (var con = Open())
            {
                var rows = con.Query(
                    @"SELECT AUTONUM,AUTONUMCNTR,LOCAL,TIPO,ID_CONTEINER,Viagem
                      FROM sgipa..VW_EGMCI_AVARIAS_TTL WHERE autonumcntr=@C",
                    new { C = autonumCntr }, commandTimeout: Config.QueryTimeoutInSeconds());
                var list = new List<AvariaCntrGridDto>();
                foreach (var row in rows)
                {
                    var d = (IDictionary<string, object>)row;
                    var aut = Convert.ToInt64(d.ContainsKey("AUTONUM") ? d["AUTONUM"] : d["autonum"]);
                    list.Add(new AvariaCntrGridDto
                    {
                        AUTONUM = aut,
                        LOCAL = AsStr(d.ContainsKey("LOCAL") ? d["LOCAL"] : d["local"]),
                        TIPO = AsStr(d.ContainsKey("TIPO") ? d["TIPO"] : d["tipo"]),
                        DESTACAR_OPERADOR = aut == 0
                    });
                }
                return list;
            }
        }

        public List<string> ListarDescricoesAvariasLocal()
        {
            using (var con = Open())
                return con.Query<string>(
                    "SELECT descr FROM sgipa..dte_tb_avarias WHERE ident='L' ORDER BY descr",
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
        }

        public List<string> ListarDescricoesAvariasTipo()
        {
            using (var con = Open())
                return con.Query<string>(
                    "SELECT descr FROM sgipa..dte_tb_avarias WHERE ident='T' ORDER BY descr",
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
        }

        public bool AvariasFinalizadas(long autonumCntr)
        {
            using (var con = Open())
            {
                var n = con.QuerySingle<int>(
                    @"SELECT COUNT(1) FROM sgipa..tb_etapas_avaria WHERE autonumcntr=@C AND cod_etapa IN (2,3,4)",
                    new { C = autonumCntr }, commandTimeout: Config.QueryTimeoutInSeconds());
                return n > 0;
            }
        }

        public void IncluirAvaria(long autonumCntr, string localDescr, string tipoDescr, int usuarioId)
        {
            using (var con = Open())
            {
                var xtipo = con.ExecuteScalar<string>(
                    "SELECT code FROM sgipa..dte_tb_avarias WHERE descr=@D AND ident='T'",
                    new { D = tipoDescr }, commandTimeout: Config.QueryTimeoutInSeconds());
                var xlocal = con.ExecuteScalar<string>(
                    "SELECT code FROM sgipa..dte_tb_avarias WHERE descr=@D AND ident='L'",
                    new { D = localDescr }, commandTimeout: Config.QueryTimeoutInSeconds());

                con.Execute(
                    @"INSERT INTO sgipa..TB_AVARIAS_CONTEINER(CNTR,LOCAL,TIPO,COMPLEMENTO,PROVIDENCIA,DT_CAD_AVARIA,USU_CAD_AVARIA)
                      SELECT @Cnt,@L,@T,3,11,GETDATE(),@U
                      WHERE NOT EXISTS (SELECT AUTONUM FROM sgipa..TB_AVARIAS_CONTEINER WHERE CNTR=@Cnt AND LOCAL=@L AND TIPO=@T)",
                    new { Cnt = autonumCntr, L = xlocal, T = xtipo, U = usuarioId },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public void ExcluirAvaria(long autonumAvaria)
        {
            using (var con = Open())
                con.Execute(
                    "DELETE FROM sgipa..TB_AVARIAS_CONTEINER WHERE AUTONUM=@A",
                    new { A = autonumAvaria },
                    commandTimeout: Config.QueryTimeoutInSeconds());
        }

        public void FinalizarAvarias(long autonumCntr)
        {
            using (var con = Open())
                con.Execute(
                    @"INSERT INTO sgipa..TB_ETAPAS_AVARIA (AUTONUMCNTR, COD_ETAPA, DT_ETAPA) VALUES (@C,2,GETDATE())",
                    new { C = autonumCntr },
                    commandTimeout: Config.QueryTimeoutInSeconds());
        }

        #endregion

        #region Helpers

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

        private static object GetCI(IDictionary<string, object> d, params string[] keys)
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

        private static int ToInt32Safe(object o, int def = 0)
        {
            if (o == null || o is DBNull) return def;
            try { return Convert.ToInt32(o, CultureInfo.InvariantCulture); }
            catch { return def; }
        }

        private static decimal ToDecimalSafe(object o)
        {
            if (o == null || o is DBNull) return 0;
            try { return Convert.ToDecimal(o, CultureInfo.InvariantCulture); }
            catch { return 0; }
        }

        private static string TrimJoin(string a, string b)
        {
            a = (a ?? "").Trim();
            b = (b ?? "").Trim();
            if (a.Length == 0) return b;
            if (b.Length == 0) return a;
            return a + " " + b;
        }

        private static string FmtData(object o)
        {
            if (o == null || o is DBNull) return "";
            DateTime dt;
            if (DateTime.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            return "";
        }

        #endregion
    }
}
