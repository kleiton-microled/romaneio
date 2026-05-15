using Dapper;
using Newtonsoft.Json.Linq;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Romaneio.Repositorios
{
    public class InventarioCargaSoltaArmazemRepositorio : IInventarioCargaSoltaArmazemRepositorio
    {
        public List<InventarioCsComboDto> ListarArmazens(int patioUsuario)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var sql = new StringBuilder();
                sql.Append("SELECT Autonum AS AUTONUM, Descr AS DISPLAY FROM SGIPA..TB_ARMAZENS_IPA WHERE ");
                if (patioUsuario != 1 && patioUsuario != 7)
                {
                    sql.Append(" PATIO = @PATIO AND (ISNULL(FLAG_CT,0)=0 OR (FLAG_CT=1 AND FLAG_HISTORICO=0)) ");
                }
                else
                {
                    sql.Append(" PATIO IN (1,7) AND (ISNULL(FLAG_CT,0)=0 OR (FLAG_CT=1 AND FLAG_HISTORICO=0)) ");
                }
                sql.Append(" ORDER BY DESCR ");

                return con.Query<InventarioCsComboDto>(
                    sql.ToString(),
                    new { PATIO = patioUsuario },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<InventarioCsComboDto> ListarMotivos()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"SELECT Autonum AS AUTONUM, Descricao AS DISPLAY
                                     FROM OPERADOR..TB_CAD_MOTIVO
                                     ORDER BY Descricao";
                return con.Query<InventarioCsComboDto>(sql, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public InventarioCsMarcanteEstadoResponse ObterEstadoMarcante(int marcante, int patioUsuario)
        {
            var res = new InventarioCsMarcanteEstadoResponse { MODO_TAG = "C", TITULO_STATUS = "INVENTARIO DE CARGA SOLTA" };
            if (marcante <= 0)
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Marcante invalido";
                return res;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                var rowFn = QueryPrimeiraLinhaComoJObject(con,
                    "SELECT * FROM SGIPA.dbo.fnInvent_Sistemas_Arm_Col_patio(@M)",
                    new { M = marcante });

                if (rowFn == null || !rowFn.Properties().Any())
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = "Marcante nao localizado";
                    return res;
                }

                int? patioMarc = GetInt(rowFn, "Patio", "PATIO");
                if (patioMarc.HasValue && patioMarc.Value != 0)
                {
                    if (patioMarc.Value != 1 && patioMarc.Value != 7)
                    {
                        if (patioMarc.Value != patioUsuario)
                        {
                            res.PATIO_DIVERGENTE = true;
                            res.SUCESSO = false;
                            res.MENSAGEM = "PATIO DIVERGENTE AO USUARIO LOGADO";
                            res.TITULO_STATUS = res.MENSAGEM;
                            return res;
                        }
                    }
                    else
                    {
                        int colPatio2 = patioUsuario == 1 ? 7 : 1;
                        if (patioUsuario != patioMarc.Value && colPatio2 != patioMarc.Value)
                        {
                            res.PATIO_DIVERGENTE = true;
                            res.SUCESSO = false;
                            res.MENSAGEM = "PATIO DIVERGENTE AO USUARIO LOGADO";
                            res.TITULO_STATUS = res.MENSAGEM;
                            return res;
                        }
                    }
                }

                int autonumCexp = GetInt(rowFn, "AUTONUM_CEXP", "autonum_cexp") ?? 0;
                int autonumCarga = GetInt(rowFn, "autonum_carga", "AUTONUM_CARGA") ?? 0;
                int volumes = GetInt(rowFn, "VOLUMES", "volumes") ?? 0;
                int autonumCsYard = GetInt(rowFn, "Autonum_Cs_Yard", "AUTONUM_CS_YARD", "autonum_cs_yard") ?? 0;
                int autonumCsYardAc = GetInt(rowFn, "Autonum_Cs_Yard_AC", "AUTONUM_CS_YARD_AC", "autonum_cs_yard_ac") ?? 0;
                int flagRecon = GetInt(rowFn, "FLAG_RECONHECIMENTO", "flag_reconhecimento") ?? 0;

                if (autonumCexp == 0 && autonumCarga > 0)
                {
                    res.MODO_TAG = "C";
                    var bl = con.QueryFirstOrDefault<long?>(
                        "SELECT BL FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM=@AC",
                        new { AC = autonumCarga },
                        commandTimeout: Config.QueryTimeoutInSeconds());
                    if (!bl.HasValue)
                    {
                        res.SUCESSO = false;
                        res.MENSAGEM = "Carga nao localizada";
                        return res;
                    }

                    PreencherPorLote(con, res, bl.Value, marcante, autonumCsYard);
                    res.QTDE = volumes;
                    res.QTDE_M = volumes;
                    res.QTDE_POS = volumes;
                    res.AUTONUM_CS_YARD = autonumCsYard;

                    if (flagRecon == 1)
                    {
                        res.FLAG_RECONHECIMENTO = true;
                        res.MODO_TAG = "M";
                        res.TITULO_STATUS = "EM RECONHECIMENTO";
                        var lotImp = GetString(rowFn, "LOTE_IMPRESSO", "lote_impresso");
                        var idCntr = GetString(rowFn, "ID_CONTEINER_IMPRESSO", "id_conteiner_impresso");
                        var armRec = GetInt(rowFn, "ARMAZEM_REC", "armazem_rec") ?? 0;
                        var yardRec = GetString(rowFn, "YARD_REC", "yard_rec");
                        res.LOTE = string.IsNullOrWhiteSpace(lotImp) ? res.LOTE : lotImp;
                        res.CONTEINER = idCntr ?? string.Empty;
                        var descrArm = con.ExecuteScalar<string>(
                            "SELECT DESCR FROM SGIPA..TB_ARMAZENS_IPA WHERE AUTONUM=@A",
                            new { A = armRec },
                            commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty;
                        res.LOCAL_ATUAL = (descrArm + " " + (yardRec ?? string.Empty)).Trim();
                    }
                }
                else if (autonumCexp == 0 && autonumCarga == 0)
                {
                    res.MODO_TAG = "Z";
                    res.TITULO_STATUS = "INVENTARIO DE CARGA SOLTA";
                    PreencherPorVwInventMarcante(con, res, marcante);
                    res.QTDE = volumes;
                    res.QTDE_M = volumes;
                    res.QTDE_POS = volumes;
                    res.AUTONUM_CS_YARD = autonumCsYardAc;
                }
                else
                {
                    res.FLAG_EXPORTACAO = true;
                    res.MODO_TAG = "C";
                    PreencherPorVwInventMarcante(con, res, marcante);
                    res.QTDE = volumes;
                    res.QTDE_M = volumes;
                    res.QTDE_POS = volumes;
                    res.AUTONUM_CS_YARD = autonumCsYard;
                    var ltExp = (res.LOCAL_ATUAL ?? string.Empty).Replace("-", string.Empty).Trim();
                    res.MOTIVO_PADRAO = string.IsNullOrEmpty(ltExp) || ltExp == "-" ? 14 : 1;
                }

                res.SUCESSO = true;
                return res;
            }
        }

        public InventarioCsMarcanteEstadoResponse ObterEstadoLote(long loteBl, int marcante, long autonumCsYard)
        {
            var res = new InventarioCsMarcanteEstadoResponse { MODO_TAG = "C" };
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                PreencherPorLote(con, res, loteBl, marcante, autonumCsYard);
                res.SUCESSO = !string.IsNullOrEmpty(res.MERCADORIA) || res.ITENS.Count > 0;
                if (!res.SUCESSO)
                {
                    res.MENSAGEM = "Lote nao encontrado";
                }
                return res;
            }
        }

        public InventarioCsItemDetalheResponse ObterDetalheItem(string idGravacao)
        {
            var res = new InventarioCsItemDetalheResponse();
            if (string.IsNullOrWhiteSpace(idGravacao))
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Item nao informado";
                return res;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"SELECT QTDE, EMBALAGEM,
                    (ISNULL(DESCR_ARMAZEM,'') + ' ' + ISNULL(POSICAO,'')) AS LOCAL
                    FROM OPERADOR..VW_INVENT_ARMAZEM
                    WHERE ID_GRAVACAO = @ID";
                var linha = con.QueryFirstOrDefault(sql, new { ID = idGravacao.Trim() },
                    commandTimeout: Config.QueryTimeoutInSeconds());
                if (linha == null)
                {
                    res.SUCESSO = false;
                    res.MENSAGEM = "Item nao encontrado";
                    return res;
                }

                var jo = JObject.FromObject(linha);
                res.QTDE = GetInt(jo, "QTDE", "qtde") ?? 0;
                res.EMBALAGEM = GetString(jo, "EMBALAGEM", "embalagem");
                res.LOCAL = GetString(jo, "LOCAL", "local");
                var localTrim = (res.LOCAL ?? string.Empty).Trim();
                res.MOTIVO_PADRAO = string.IsNullOrEmpty(localTrim) || localTrim == "-" ? 14 : 1;
                res.SUCESSO = true;
                return res;
            }
        }

        public List<InventarioCsHistoricoLinhaDto> ListarHistorico(int marcante)
        {
            if (marcante <= 0)
            {
                return new List<InventarioCsHistoricoLinhaDto>();
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"
SELECT ISNULL(B.DESCR,'') AS ARMAZEM,
       UPPER(ISNULL(A.YARD,'')) AS LOCAL,
       RTRIM(LTRIM(CONVERT(CHAR(10), A.DT_MOV, 3))) + ' ' + LEFT(RTRIM(LTRIM(CONVERT(CHAR(12), A.DT_MOV, 108))), 5) AS DATA_MOV,
       U.USUARIO AS USUARIO
FROM SGIPA..TB_HIST_SHIFTING_CS A
LEFT JOIN SGIPA..TB_ARMAZENS_IPA B ON A.ARMAZEM = B.AUTONUM
LEFT JOIN SGIPA..TB_CAD_USUARIOS U ON A.USUARIO = U.AUTONUM
WHERE A.MARCANTE = @M
ORDER BY A.DT_MOV DESC";

                return con.Query<InventarioCsHistoricoLinhaDto>(sql, new { M = marcante },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public InventarioCsResolverPrateleiraResponse ResolverPrateleira(string texto, int patioUsuario)
        {
            var res = new InventarioCsResolverPrateleiraResponse();
            if (string.IsNullOrWhiteSpace(texto))
            {
                res.MENSAGEM = "Etiqueta vazia";
                return res;
            }

            texto = texto.Trim();
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();

                if (texto.Length == 12 && texto.IndexOf('-') == 10)
                {
                    var tbCt = con.QueryFirstOrDefault<int?>(
                        @"SELECT AUTONUM FROM SGIPA..TB_ARMAZENS_IPA
                          WHERE DESCR=@D AND ISNULL(FLAG_CT,0)=1 AND ISNULL(FLAG_HISTORICO,0)=0",
                        new { D = texto },
                        commandTimeout: Config.QueryTimeoutInSeconds());
                    if (tbCt.HasValue)
                    {
                        res.SUCESSO = true;
                        res.MODO_CONTENER = true;
                        res.ARMAZEM_ID = tbCt.Value;
                        res.YARD = string.Empty;
                        return res;
                    }
                }

                string autonumArmStr;
                string yardCs;
                if (texto.Length == 9 || texto.Length == 10 || (texto.Length >= 4 && texto.Length < 12 && texto.IndexOf('-') < 0) ||
                    (texto.Length == 12 && texto.IndexOf('-') == 3))
                {
                    if (patioUsuario == 1 || patioUsuario == 7)
                    {
                        autonumArmStr = texto.Substring(0, 2);
                    }
                    else
                    {
                        var dois = texto.Substring(0, 2);
                        autonumArmStr = con.ExecuteScalar<string>(
                            @"SELECT CONVERT(VARCHAR(20), MAX(AUTONUM)) FROM SGIPA..TB_ARMAZENS_IPA
                              WHERE PATIO=@P AND AUTONUM_ARM_PATIO=@AP",
                            new { P = patioUsuario, AP = int.Parse(dois, CultureInfo.InvariantCulture) },
                            commandTimeout: Config.QueryTimeoutInSeconds()) ?? "0";
                    }

                    yardCs = texto.Length > 2
                        ? texto.Substring(2).Replace("-", string.Empty).ToUpperInvariant()
                        : string.Empty;

                    yardCs = AjustarYardViaTbYardCs(con, autonumArmStr, yardCs, patioUsuario, texto.Length);
                    if (string.IsNullOrEmpty(yardCs))
                    {
                        res.MENSAGEM = "Posicao nao encontrada";
                        return res;
                    }

                    res.SUCESSO = true;
                    res.ARMAZEM_ID = int.Parse(autonumArmStr, CultureInfo.InvariantCulture);
                    res.YARD = yardCs;
                    return res;
                }

                res.MENSAGEM = "Formato de etiqueta nao reconhecido";
                return res;
            }
        }

        public InventarioCsSalvarResult Salvar(InventarioCsSalvarRequest request, int codUsuario, int patioUsuario)
        {
            var result = new InventarioCsSalvarResult();
            if (request == null)
            {
                result.MENSAGEM = "Requisicao invalida";
                return result;
            }

            string erro = ValidarSalvar(request, patioUsuario);
            if (!string.IsNullOrEmpty(erro))
            {
                result.MENSAGEM = erro;
                return result;
            }

            var mov = CalcularParametrosMovimentacao(request);
            int patioProc = patioUsuario == 7 ? 1 : patioUsuario;
            string yardProc = request.YARD ?? string.Empty;

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                try
                {
                    var p = new DynamicParameters();
                    p.Add("Tag", mov.Tag, DbType.AnsiStringFixedLength, size: 1);
                    p.Add("Tipo", mov.Tipo, DbType.AnsiStringFixedLength, size: 1);
                    p.Add("FlagChronos", mov.FlagChronos ? 1 : 0, DbType.Int32);
                    p.Add("AutonumCs", mov.AutonumCs, DbType.Int32);
                    p.Add("Patio", patioProc, DbType.Int32);
                    p.Add("Yard", yardProc.Length > 10 ? yardProc.Substring(0, 10) : yardProc, DbType.AnsiString, size: 10);
                    p.Add("Quantidade", request.QUANTIDADE_POS, DbType.Int32);
                    p.Add("Motivo", request.MOTIVO_ID, DbType.Int32);
                    p.Add("Marcante", request.MARCANTE, DbType.Int32);
                    p.Add("Usuario", codUsuario, DbType.Int32);
                    p.Add("FlagExp", request.FLAG_EXPORTACAO ? 1 : 0, DbType.Int32);
                    p.Add("Armazem", request.ARMAZEM_ID, DbType.Int32);
                    p.Add("Fl_Frente", request.FL_FRENTE, DbType.Int32);
                    p.Add("Fl_Fundo", request.FL_FUNDO, DbType.Int32);
                    p.Add("Fl_LE", request.FL_LE, DbType.Int32);
                    p.Add("Fl_LD", request.FL_LD, DbType.Int32);
                    p.Add("Autonumcy", mov.AutonumCY, DbType.Int32);
                    var yardAtual = request.YARD_ATUAL_TEXTO ?? string.Empty;
                    if (yardAtual.Length > 30)
                    {
                        yardAtual = yardAtual.Substring(0, 30);
                    }
                    p.Add("YardAtual", yardAtual, DbType.AnsiString, size: 30);
                    p.Add("FlagCT", request.FLAG_MODO_CONTENER ? 1 : 0, DbType.Int32);
                    p.Add("PercentCT", request.PERCENT_CT, DbType.Int32);

                    con.Execute(
                        "Sgipa.dbo.pcMovimentaCS2",
                        p,
                        commandTimeout: Config.QueryTimeoutInSeconds(),
                        commandType: CommandType.StoredProcedure);
                }
                catch (SqlException ex)
                {
                    result.MENSAGEM = "Erro ao executar movimentacao: " + ex.Message;
                    return result;
                }

                // Pos-processo segregacao IMO (legado VB chama Verifica_Regras_Seg_Imo_DeltaCS - modulo nao disponivel no repo).
                // Pendencia tecnica: portar equivalencia quando existir funcao SQL ou servico compartilhado.
                // Pos-processo segregacao IMO: legado chama Verifica_Regras_Seg_Imo_DeltaCS (modulo VB indisponivel aqui).
                // Pendencia: integrar quando existir funcao SQL equivalente ou servico compartilhado.

                result.SUCESSO = true;
                result.MENSAGEM = "CARGA POSICIONADA COM SUCESSO";
                return result;
            }
        }

        private class MovimentacaoParams
        {
            public string Tipo { get; set; }
            public int AutonumCs { get; set; }
            public int AutonumCY { get; set; }
            public bool FlagChronos { get; set; }
            public string Tag { get; set; }
        }

        private MovimentacaoParams CalcularParametrosMovimentacao(InventarioCsSalvarRequest r)
        {
            var m = new MovimentacaoParams
            {
                Tag = string.IsNullOrEmpty(r.MODO_TAG) ? "C" : r.MODO_TAG.Substring(0, 1),
                FlagChronos = r.MODO_TAG == "C" || r.MODO_TAG == "M"
            };

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                if (!r.FLAG_EXPORTACAO)
                {
                    if (m.FlagChronos)
                    {
                        var chave = r.ITEM_CHAVE ?? string.Empty;
                        m.Tipo = chave.Length > 0 ? chave.Substring(0, 1).ToUpperInvariant() : "C";
                        if (m.Tipo == "C")
                        {
                            m.AutonumCs = ParseId(chave, 1);
                        }
                        else if (m.Tipo == "Y")
                        {
                            m.AutonumCY = ParseId(chave, 1);
                            m.AutonumCs = con.ExecuteScalar<int?>(
                                "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD WHERE AUTONUM=@Y",
                                new { Y = m.AutonumCY },
                                commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
                        }
                    }
                    else
                    {
                        var tag = r.LOTE_TAG ?? string.Empty;
                        m.Tipo = tag.Length > 0 ? tag.Substring(0, 1).ToUpperInvariant() : "V";
                        if (m.Tipo == "V")
                        {
                            m.AutonumCs = ParseId(tag, 1);
                        }
                        else if (m.Tipo == "W")
                        {
                            m.AutonumCY = ParseId(tag, 1);
                            m.AutonumCs = con.ExecuteScalar<int?>(
                                "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD_AC WHERE AUTONUM=@Y",
                                new { Y = m.AutonumCY },
                                commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
                        }
                    }
                }
                else
                {
                    var tag = r.LOTE_TAG ?? string.Empty;
                    m.Tipo = tag.Length > 0 ? tag.Substring(0, 1).ToUpperInvariant() : "R";
                    if (m.Tipo == "R")
                    {
                        m.AutonumCs = ParseId(tag, 1);
                    }
                    else if (m.Tipo == "T")
                    {
                        m.AutonumCY = ParseId(tag, 1);
                        m.AutonumCs = con.ExecuteScalar<int?>(
                            "SELECT AUTONUM_CEXP FROM SGIPA..TB_CARGA_SOLTA_YARD_CEXP WHERE AUTONUM=@Y",
                            new { Y = m.AutonumCY },
                            commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
                    }
                }
            }

            return m;
        }

        private static int ParseId(string s, int startIndex)
        {
            if (string.IsNullOrEmpty(s) || startIndex >= s.Length)
            {
                return 0;
            }

            int v;
            return int.TryParse(s.Substring(startIndex), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? v : 0;
        }

        private string ValidarSalvar(InventarioCsSalvarRequest r, int patioUsuario)
        {
            if (string.IsNullOrWhiteSpace(r.LOTE))
            {
                return "INFORME O LOTE";
            }

            if (r.MODO_TAG == "C" && !r.FLAG_EXPORTACAO && string.IsNullOrWhiteSpace(r.ITEM_CHAVE))
            {
                return "ITEM NAO INFORMADO";
            }

            if (r.QUANTIDADE_POS <= 0)
            {
                return "QUANTIDADE NAO INFORMADA";
            }

            if (r.QUANTIDADE_POS > r.QTDE_DISPONIVEL)
            {
                return "QUANTIDADE INDISPONIVEL";
            }

            if (r.ARMAZEM_ID <= 0)
            {
                return "ARMAZEM NAO INFORMADO";
            }

            if (!r.FLAG_MODO_CONTENER && string.IsNullOrWhiteSpace(r.YARD))
            {
                return "LOCAL NAO INFORMADO";
            }

            if (!r.FLAG_EXPORTACAO && string.Equals(r.YARD, "CAM", StringComparison.OrdinalIgnoreCase))
            {
                return "CARREGAMENTO DO IPA INVALIDO";
            }

            if (r.MOTIVO_ID <= 0)
            {
                return "MOTIVO NAO INFORMADO";
            }

            if (r.ARMAZEM_ID == 22)
            {
                using (SqlConnection con = new SqlConnection(Config.StringConexao()))
                {
                    var imo = con.ExecuteScalar<string>(
                        "SELECT IMO FROM SGIPA..TB_CARGA_SOLTA WHERE BL=@BL AND ISNULL(IMO,'')<>''",
                        new { BL = r.LOTE },
                        commandTimeout: Config.QueryTimeoutInSeconds());
                    if (!string.IsNullOrWhiteSpace(imo))
                    {
                        var up = (r.YARD ?? string.Empty).ToUpperInvariant();
                        var infl = new[] { "2.1", "3", "4.1", "4.2", "4.3", "3.0" };
                        if (infl.Contains(imo))
                        {
                            if (imo != "4.3")
                            {
                                if (!ContainsL1L10(up))
                                {
                                    return "Local Bloqueado, carga Inflamavel";
                                }
                            }
                            else if (up != "L10")
                            {
                                return "Local Bloqueado, carga Inflamavel (L10)";
                            }
                        }
                    }
                }
            }

            if (r.MARCANTE > 0)
            {
                using (SqlConnection con = new SqlConnection(Config.StringConexao()))
                {
                    var tbM = con.QueryFirstOrDefault(
                        "SELECT AUTONUM, AUTONUM_CARGA, ISNULL(AUTONUM_CARGA_AC,0) AS AUTONUM_CARGA_AC FROM SGIPA..TB_MARCANTES WHERE AUTONUM=@M",
                        new { M = r.MARCANTE },
                        commandTimeout: Config.QueryTimeoutInSeconds());
                    if (tbM == null)
                    {
                        return "MARCANTE INVALIDO";
                    }

                    var d = (IDictionary<string, object>)tbM;
                    int cargAc = ToInt(GetFromDict(d, "AUTONUM_CARGA_AC"));
                    if (cargAc == 0)
                    {
                        var chave = r.ITEM_CHAVE ?? string.Empty;
                        var tipo = chave.Length > 0 ? chave.Substring(0, 1).ToUpperInvariant() : string.Empty;
                        int autonumCs = 0;
                        if (tipo == "C")
                        {
                            autonumCs = ParseId(chave, 1);
                        }
                        else if (tipo == "Y")
                        {
                            var cy = ParseId(chave, 1);
                            autonumCs = con.ExecuteScalar<int?>(
                                "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD WHERE AUTONUM=@Y",
                                new { Y = cy },
                                commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;
                        }

                        int autonumCarga = ToInt(GetFromDict(d, "AUTONUM_CARGA"));
                        if (autonumCarga > 0 && autonumCarga != autonumCs)
                        {
                            return "MARCANTE ASSOCIADO A OUTRA CARGA";
                        }
                    }
                }
            }

            if (r.QTDE_MARCANTE > 0 && r.QUANTIDADE_POS != r.QTDE_MARCANTE)
            {
                return "DIVERGENCIA COM A QTDE MARCANTE";
            }

            if (!r.FLAG_MODO_CONTENER)
            {
                var yardInformado = r.YARD ?? string.Empty;
                string localCorrigido = null;
                if (yardInformado.Length == 4)
                {
                    using (SqlConnection con = new SqlConnection(Config.StringConexao()))
                    {
                        var ex = con.ExecuteScalar<int?>(
                            "SELECT AUTONUM FROM SGIPA..TB_YARD_CS WHERE YARD=@Y",
                            new { Y = yardInformado },
                            commandTimeout: Config.QueryTimeoutInSeconds());
                        if (!ex.HasValue)
                        {
                            localCorrigido = yardInformado.Substring(0, 3) + "01" + yardInformado.Substring(3, 1);
                        }
                    }
                }

                using (SqlConnection con = new SqlConnection(Config.StringConexao()))
                {
                    object rst = con.QueryFirstOrDefault(
                        @"SELECT AUTONUM, VALIDA, FLAG_BLOQUEIO, LIMITE_MARCANTES, YARD
                          FROM SGIPA..TB_YARD_CS
                          WHERE ARMAZEM=@A AND YARD=@Y",
                        new { A = r.ARMAZEM_ID, Y = yardInformado },
                        commandTimeout: Config.QueryTimeoutInSeconds());

                    if (rst == null && !string.IsNullOrEmpty(localCorrigido))
                    {
                        rst = con.QueryFirstOrDefault(
                            @"SELECT AUTONUM, VALIDA, FLAG_BLOQUEIO, LIMITE_MARCANTES, YARD
                              FROM SGIPA..TB_YARD_CS
                              WHERE ARMAZEM=@A AND YARD=@Y",
                            new { A = r.ARMAZEM_ID, Y = localCorrigido },
                            commandTimeout: Config.QueryTimeoutInSeconds());
                    }

                    if (rst == null)
                    {
                        return "LOCAL NAO ENCONTRADO";
                    }

                    var d = (IDictionary<string, object>)rst;
                    if (ToInt(GetFromDict(d, "FLAG_BLOQUEIO")) == 1)
                    {
                        return "LOCAL BLOQUEADO";
                    }

                    if (ToInt(GetFromDict(d, "VALIDA")) != 1)
                    {
                        int lim = ToInt(GetFromDict(d, "LIMITE_MARCANTES"));
                        if (lim > 0)
                        {
                            var yardRow = GetFromDict(d, "YARD");
                            int qtos = con.ExecuteScalar<int>(
                                @"SELECT COUNT(0) FROM OPERADOR..VW_INVENT_ARMAZEM
                                  WHERE AUTONUM_ARMAZEM=@A AND POSICAO=@P",
                                new { A = r.ARMAZEM_ID, P = yardRow },
                                commandTimeout: Config.QueryTimeoutInSeconds());
                            if (qtos > lim)
                            {
                                return "LIMITE DE CARGA NO LOCAL";
                            }
                        }
                    }

                    if (ToInt(GetFromDict(d, "FLAG_BLOQUEIO")) == 1)
                    {
                        return "LOCAL BLOQUEADO";
                    }
                }
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var retAnv = con.ExecuteScalar<string>(
                    "SELECT [dbo].[FN_VALIDA_ANVISA_ARMAZEM](@L,@A,@Y)",
                    new { L = r.LOTE, A = r.ARMAZEM_ID, Y = r.YARD ?? string.Empty },
                    commandTimeout: Config.QueryTimeoutInSeconds());
                if (!string.Equals(retAnv, "OK", StringComparison.OrdinalIgnoreCase))
                {
                    return retAnv ?? "Validacao ANVISA falhou";
                }

                var retLote = con.ExecuteScalar<string>(
                    "SELECT [OPERADOR].[dbo].[FN_VALIDA_LOTE_PRATELEIRA](@L,@A,@Y)",
                    new { L = r.LOTE, A = r.ARMAZEM_ID, Y = r.YARD ?? string.Empty },
                    commandTimeout: Config.QueryTimeoutInSeconds());
                if (!string.Equals(retLote, "OK", StringComparison.OrdinalIgnoreCase))
                {
                    return "Outro Lote na posicao " + (retLote ?? string.Empty);
                }
            }

            return null;
        }

        private static bool ContainsL1L10(string up)
        {
            return up == "L1" || up == "L2" || up == "L3" || up == "L4" || up == "L5"
                || up == "L6" || up == "L7" || up == "L8" || up == "L9" || up == "L10";
        }

        private void PreencherPorLote(SqlConnection con, InventarioCsMarcanteEstadoResponse res, long loteBl, int marcante, long autonumCsYard)
        {
            var row = QueryPrimeiraLinhaComoJObject(con,
                "SELECT * FROM SGIPA.dbo.fnInvent_Sistemas_Arm_CollOTE(@L,@M)",
                new { L = loteBl, M = marcante });

            if (row == null || !row.Properties().Any())
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Lote nao encontrado";
                return;
            }

            res.MERCADORIA = GetString(row, "MERCADORIA", "mercadoria");
            res.MARCA = GetString(row, "Marca", "MARCA", "marca");
            res.ENTRADA = GetString(row, "Entrada", "ENTRADA", "data_entrada");
            res.CONTEINER = GetString(row, "CNTR_DESOVA", "cntr_desova");
            res.CLIENTE = GetString(row, "Importador", "IMPORTADOR", "importador");
            var tipoDoc = GetString(row, "Tipo_Doc", "TIPO_DOC", "tipo_doc");
            res.DOC = string.Equals(tipoDoc, "NOT DEFINED YET", StringComparison.OrdinalIgnoreCase) ? string.Empty : tipoDoc;
            res.CANAL = GetString(row, "CANAL_ALF", "canal_alf");
            res.MOV = GetString(row, "MOTIVO_PROX_MVTO", "motivo_prox_mvto");
            var imo = GetString(row, "IMO", "imo");
            res.IMO = imo == "0" ? string.Empty : imo;
            res.NVOCC = GetString(row, "NVOCC", "nvocc");
            res.BL = GetString(row, "BL", "bl");
            var descr = GetString(row, "DESCR_ARMAZEM", "descr_armazem");
            var pos = GetString(row, "Posicao", "POSICAO", "posicao");
            res.LOCAL_ATUAL = (descr + " " + pos).Trim();
            var localNorm = res.LOCAL_ATUAL.Replace("-", string.Empty).Trim();
            if (string.IsNullOrEmpty(localNorm) || localNorm == "")
                res.MOTIVO_PADRAO = 14;
            else
                res.MOTIVO_PADRAO = 1;

            res.ITENS = CarregarItens(con, loteBl, marcante, autonumCsYard);
            res.SUCESSO = true;
        }

        private void PreencherPorVwInventMarcante(SqlConnection con, InventarioCsMarcanteEstadoResponse res, int marcante)
        {
            const string sql = @"SELECT LOTE, BL, EMBALAGEM, IMPORTADOR, NAVIO, QTDE, AUTONUMCS, MERCADORIA, DATA_ENTRADA,
                                        TIPO_DOC, DESCR_ARMAZEM, POSICAO, ID_GRAVACAO
                                 FROM OPERADOR..VW_INVENT_ARMAZEM WHERE MARCANTE = @M";
            var linha = con.QueryFirstOrDefault(sql, new { M = marcante }, commandTimeout: Config.QueryTimeoutInSeconds());
            if (linha == null)
            {
                res.SUCESSO = false;
                res.MENSAGEM = "Marcante sem inventario";
                return;
            }

            var d = (IDictionary<string, object>)linha;
            res.LOTE = GetFromDict(d, "LOTE");
            res.LOTE_TAG = GetFromDict(d, "ID_GRAVACAO");
            res.MERCADORIA = GetFromDict(d, "MERCADORIA");
            res.MARCA = string.Empty;
            res.ENTRADA = GetFromDict(d, "DATA_ENTRADA");
            res.CONTEINER = string.Empty;
            res.CLIENTE = GetFromDict(d, "IMPORTADOR");
            res.EMBALAGEM = GetFromDict(d, "EMBALAGEM");
            var tipoDoc = GetFromDict(d, "TIPO_DOC");
            res.DOC = string.Equals(tipoDoc, "NOT DEFINED YET", StringComparison.OrdinalIgnoreCase) ? string.Empty : tipoDoc;
            res.CANAL = string.Empty;
            res.MOV = string.Empty;
            res.IMO = string.Empty;
            res.NVOCC = string.Empty;
            res.BL = GetFromDict(d, "BL");
            res.LOCAL_ATUAL = (GetFromDict(d, "DESCR_ARMAZEM") + " " + GetFromDict(d, "POSICAO")).Trim();
            var lt = res.LOCAL_ATUAL.Replace("-", string.Empty).Trim();
            res.MOTIVO_PADRAO = string.IsNullOrEmpty(lt) || lt == "-" ? 14 : 1;
        }

        private List<InventarioCsItemLinhaDto> CarregarItens(SqlConnection con, long loteBl, int marcante, long autonumCsYard)
        {
            var lista = new List<InventarioCsItemLinhaDto>();
            var rows = con.Query(
                "SELECT * FROM SGIPA.dbo.fnInvent_Sistemas_Arm_ColItem(@L,@M,@Y)",
                new { L = loteBl, M = marcante, Y = autonumCsYard },
                commandTimeout: Config.QueryTimeoutInSeconds()).ToList();

            foreach (var row in rows)
            {
                var jo = JObject.FromObject(row);
                string chave = GetString(jo, "autonum", "AUTONUM", "ID_GRAVACAO", "id_gravacao", "ITEM_CHAVE");
                if (string.IsNullOrWhiteSpace(chave))
                {
                    continue;
                }

                string display = MontarDisplayItem(jo);
                lista.Add(new InventarioCsItemLinhaDto { ITEM_CHAVE = chave.Trim(), DISPLAY = display });
            }

            return lista;
        }

        private static string MontarDisplayItem(JObject jo)
        {
            var sb = new StringBuilder();
            foreach (var p in jo.Properties())
            {
                if (p.Value == null || p.Value.Type == JTokenType.Null)
                {
                    continue;
                }

                sb.Append(p.Name).Append('=').Append(p.Value).Append(' ');
            }

            return sb.ToString().Trim();
        }

        private string AjustarYardViaTbYardCs(SqlConnection con, string autonumArmStr, string yardCs, int patioUsuario, int lenOriginal)
        {
            int arm = int.Parse(autonumArmStr, CultureInfo.InvariantCulture);
            var xpos = yardCs;
            var existe = con.ExecuteScalar<int?>(
                "SELECT AUTONUM FROM SGIPA..TB_YARD_CS WHERE ARMAZEM=@A AND YARD=@Y",
                new { A = arm, Y = xpos },
                commandTimeout: Config.QueryTimeoutInSeconds());
            if (existe.HasValue)
            {
                return yardCs;
            }

            string candidato;
            if (patioUsuario == 1 || patioUsuario == 7)
            {
                if (yardCs.Length >= 4)
                {
                    candidato = yardCs.Substring(0, 3) + "01" + yardCs.Substring(yardCs.Length - 1, 1);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                if (yardCs.Length >= 5)
                {
                    candidato = yardCs.Substring(0, 4) + "01" + yardCs.Substring(yardCs.Length - 1, 1);
                }
                else
                {
                    return string.Empty;
                }
            }

            var ex2 = con.ExecuteScalar<int?>(
                "SELECT AUTONUM FROM SGIPA..TB_YARD_CS WHERE ARMAZEM=@A AND YARD=@Y",
                new { A = arm, Y = candidato },
                commandTimeout: Config.QueryTimeoutInSeconds());
            return ex2.HasValue ? candidato : string.Empty;
        }

        private static JObject QueryPrimeiraLinhaComoJObject(SqlConnection con, string sql, object param)
        {
            var dyn = con.Query(sql, param, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            if (dyn == null)
            {
                return null;
            }

            return JObject.FromObject(dyn);
        }

        private static int? GetInt(JObject o, params string[] names)
        {
            foreach (var n in names)
            {
                var t = o.GetValue(n, StringComparison.OrdinalIgnoreCase);
                if (t != null && t.Type != JTokenType.Null)
                {
                    int v;
                    if (int.TryParse(t.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out v))
                    {
                        return v;
                    }
                }
            }

            return null;
        }

        private static string GetString(JObject o, params string[] names)
        {
            foreach (var n in names)
            {
                var t = o.GetValue(n, StringComparison.OrdinalIgnoreCase);
                if (t != null && t.Type != JTokenType.Null)
                {
                    return t.ToString();
                }
            }

            return string.Empty;
        }

        private static string GetFromDict(IDictionary<string, object> d, string key)
        {
            foreach (var kv in d)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    return kv.Value == null ? string.Empty : Convert.ToString(kv.Value, CultureInfo.InvariantCulture);
                }
            }

            return string.Empty;
        }

        private static int ToInt(object o)
        {
            if (o == null)
            {
                return 0;
            }

            int v;
            return int.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out v) ? v : 0;
        }
    }
}
