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
    public class CarregamentoCargaSoltaRepositorio : ICarregamentoCargaSoltaRepositorio
    {
        public List<VeiculoCarregamentoDto> ListarVeiculos(int patioUsuario)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                string filtroPatio;
                DynamicParameters parametros = new DynamicParameters();

                // Regra do legado (CarregaCS.frm - Carrega_Veiculos):
                // Quando Col_Patio = 1, expande para PATIO IN (1,7); caso contrario, exato.
                if (patioUsuario == 1 || patioUsuario == 7)
                {
                    filtroPatio = "PATIO IN (1, 7)";
                }
                else
                {
                    filtroPatio = "PATIO = @PATIO";
                    parametros.Add("PATIO", patioUsuario);
                }

                string sql = @"SELECT DISTINCT
                                   PLACA_C,
                                   PLACA_CARRETA,
                                   ISNULL(MODELO, '') AS MODELO,
                                   PLACA_C + ' ' + PLACA_CARRETA + ' - ' + ISNULL(MODELO, '') AS DISPLAY
                               FROM SGIPA..VW_CAM_CARREGAMENTO
                               WHERE " + filtroPatio + @"
                               ORDER BY DISPLAY";

                return con.Query<VeiculoCarregamentoDto>(sql, parametros, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<OrdemCarregamentoDto> ListarOrdens(string placaCavalo)
        {
            if (string.IsNullOrWhiteSpace(placaCavalo))
            {
                return new List<OrdemCarregamentoDto>();
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"
                    SELECT
                        ISNULL(A.NUM_OC, 0)        AS NUM_OC,
                        ISNULL(A.LOTE, 0)          AS LOTE,
                        ISNULL(A.ITEM, 0)          AS ITEM,
                        ISNULL(A.QUANTIDADE, 0)    AS QUANTIDADE,
                        ISNULL(B.QTDE_CARREGADA,0) AS QTDE_CARREGADA,
                        ISNULL(A.EMBALAGEM, '')    AS EMBALAGEM,
                        ISNULL(A.ORDEM_CARREG, 0)  AS ORDEM_CARREG,
                        ISNULL(A.AUTONUMCS, 0)     AS AUTONUMCS
                    FROM SGIPA..VW_CAM_CARREGAMENTO A
                    LEFT JOIN (
                        SELECT SUM(M.VOLUMES)        AS QTDE_CARREGADA,
                               M.AUTONUM_CARGA       AS AUTONUMCS
                        FROM   SGIPA..TB_MARCANTES M
                        INNER JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD = Y.AUTONUM
                        WHERE  M.VOLUMES > 0
                          AND  Y.YARD = 'CAM'
                          AND (M.PLACA_C IS NULL OR M.PLACA_C = @PLACA)
                        GROUP BY M.AUTONUM_CARGA
                    ) B ON A.AUTONUMCS = B.AUTONUMCS
                    WHERE A.PLACA_C = @PLACA
                    ORDER BY A.LOTE, A.ITEM";

                return con.Query<OrdemCarregamentoDto>(
                    sql,
                    new { PLACA = placaCavalo },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<MarcanteCarregadoDto> ListarCarregamentos(string placaCavalo, int autonumCs)
        {
            if (autonumCs <= 0 || string.IsNullOrWhiteSpace(placaCavalo))
            {
                return new List<MarcanteCarregadoDto>();
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"
                    SELECT
                        M.AUTONUM        AS MARCANTE,
                        ISNULL(M.VOLUMES, 0) AS QTDE
                    FROM SGIPA..TB_MARCANTES M
                    INNER JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD = Y.AUTONUM
                    WHERE M.AUTONUM_CARGA = @AUTONUMCS
                      AND (M.PLACA_C IS NULL OR M.PLACA_C = @PLACA)
                      AND Y.YARD = 'CAM'
                    ORDER BY M.AUTONUM";

                return con.Query<MarcanteCarregadoDto>(
                    sql,
                    new { AUTONUMCS = autonumCs, PLACA = placaCavalo },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<LocalArmazemDto> ListarLocais(int autonumCs)
        {
            if (autonumCs <= 0)
            {
                return new List<LocalArmazemDto>();
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"
                    SELECT
                        MARCANTE,
                        ISNULL(QTDE, 0)         AS QTDE,
                        ISNULL(DESCR_ARMAZEM,'') AS DESCR_ARMAZEM,
                        ISNULL(POSICAO,'')       AS POSICAO
                    FROM SGIPA..VW_INVENT_ARMAZEM
                    WHERE AUTONUMCS = @AUTONUMCS
                      AND QTDE > 0
                    ORDER BY DESCR_ARMAZEM, POSICAO, MARCANTE";

                return con.Query<LocalArmazemDto>(
                    sql,
                    new { AUTONUMCS = autonumCs },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public MarcanteDadosDto BuscarMarcante(int marcante)
        {
            if (marcante <= 0)
            {
                return null;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"
                    SELECT
                        M.AUTONUM                       AS MARCANTE,
                        ISNULL(S.AUTONUM, 0)            AS AUTONUMCS,
                        ISNULL(S.BL, 0)                 AS LOTE,
                        ISNULL(S.ITEM, 0)               AS ITEM,
                        ISNULL(M.VOLUMES, 0)            AS VOLUMES,
                        ISNULL(E.DESCR, '')             AS EMBALAGEM,
                        ISNULL(S.MERCADORIA, '')        AS MERCADORIA,
                        ISNULL(S.MARCA, '')             AS MARCA,
                        ISNULL(S.ARMAZEM_IPA, 0)        AS AUTONUM_ARMAZEM,
                        ISNULL(A.DESCR, '')             AS DESCR_ARMAZEM,
                        S.CNTR                          AS AUTONUMCNTR,
                        ISNULL(C.Id_Conteiner, '')      AS ID_CONTEINER,
                        ISNULL(Y.YARD, '')              AS POSICAO_YARD,
                        ISNULL(M.AUTONUM_CS_YARD, 0)    AS AUTONUM_CS_YARD
                    FROM   SGIPA..TB_CARGA_SOLTA       S
                    INNER JOIN SGIPA..DTE_TB_EMBALAGENS E ON S.EMBALAGEM = E.CODE
                    LEFT  JOIN SGIPA..TB_CNTR_BL        C ON S.CNTR = C.AUTONUM
                    INNER JOIN SGIPA..TB_MARCANTES      M ON S.AUTONUM = M.AUTONUM_CARGA
                    LEFT  JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD = Y.AUTONUM
                    LEFT  JOIN SGIPA..TB_ARMAZENS_IPA   A ON Y.ARMAZEM = A.AUTONUM
                    WHERE  M.AUTONUM = @MARCANTE";

                return con.QueryFirstOrDefault<MarcanteDadosDto>(
                    sql,
                    new { MARCANTE = marcante },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public CarregarItemResult Carregar(CarregarItemRequest request, int codUsuario)
        {
            CarregarItemResult preValidacao = ValidarRequestCarregar(request);
            if (preValidacao != null)
            {
                return preValidacao;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();

                // 1) Recupera estado atual do marcante (volumes original + yard atual + autonumcs)
                MarcanteEstadoDto estado = ObterEstadoMarcante(con, request.MARCANTE, null);
                if (estado == null)
                {
                    return Falha("Marcante nao informado/invalido");
                }

                if (estado.AUTONUMCS != request.AUTONUMCS_OC)
                {
                    return Falha("Carga Divergente da Ordem de carregamento");
                }

                if (string.Equals(estado.YARD, "CAM", StringComparison.OrdinalIgnoreCase))
                {
                    return Falha("Carga ja carregada");
                }

                if (request.QTDE > estado.VOLUMES)
                {
                    return Falha("Quantidade indisponivel");
                }

                // 2) Recupera BL da carga e chama FC_VALIDA_SAIDA_CARGA antes de iniciar transacao.
                int bl = con.QueryFirstOrDefault<int?>(
                    "SELECT BL FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM = @AUTONUMCS",
                    new { AUTONUMCS = request.AUTONUMCS_OC },
                    commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;

                if (bl <= 0)
                {
                    return Falha("Lote nao encontrado");
                }

                string criticaSaida = con.ExecuteScalar<string>(
                    "SELECT dbo.FC_VALIDA_SAIDA_CARGA(@BL, @ORDEM, @AUTONUMCS) AS RETORNO",
                    new
                    {
                        BL = bl,
                        ORDEM = request.ORDEM,
                        AUTONUMCS = request.AUTONUMCS_OC
                    },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                if (!string.Equals(criticaSaida, "OK", StringComparison.OrdinalIgnoreCase))
                {
                    return Falha(criticaSaida ?? "Carregamento bloqueado");
                }

                // 3) Operacao transacional: associa marcante, atualiza yard, religa marcante e grava historico.
                using (IDbTransaction tx = con.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        // Decisao de migracao: trazer este UPDATE para dentro da transacao
                        // (no legado VB6, este UPDATE ocorria fora do BeginTrans, abrindo janela
                        // de inconsistencia caso o INSERT seguinte falhasse).
                        con.Execute(
                            @"UPDATE SGIPA..TB_MARCANTES
                              SET ORDEM = @ORDEM,
                                  PLACA_C = @PLACA
                              WHERE AUTONUM = @MARCANTE",
                            new
                            {
                                ORDEM = request.ORDEM,
                                PLACA = LimitarPlaca(request.PLACA_C),
                                MARCANTE = request.MARCANTE
                            },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        bool parcial = request.QTDE < estado.VOLUMES;

                        // 3.1) Atualiza o registro yard atual (zera ou decrementa).
                        if (!parcial)
                        {
                            con.Execute(
                                "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE = 0 WHERE AUTONUM = @ID",
                                new { ID = request.AUTONUM_CS_YARD },
                                tx,
                                Config.QueryTimeoutInSeconds());
                        }
                        else
                        {
                            con.Execute(
                                "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE = QUANTIDADE - @Q WHERE AUTONUM = @ID",
                                new { Q = request.QTDE, ID = request.AUTONUM_CS_YARD },
                                tx,
                                Config.QueryTimeoutInSeconds());
                        }

                        // 3.2) Insere novo registro yard CAM e captura o autonum gerado de forma atomica.
                        // Decisao de migracao: usar OUTPUT INSERTED em vez de SELECT MAX(AUTONUM)
                        // para corrigir risco de troca de IDs sob concorrencia.
                        int novoIdYard = con.QuerySingle<int>(
                            @"INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD
                                  (AUTONUM_CS, ARMAZEM, YARD, ORIGEM, QUANTIDADE, MOTIVO)
                              OUTPUT INSERTED.AUTONUM
                              VALUES (@AUTONUMCS, @ARMAZEM, 'CAM', 'I', @Q, 8)",
                            new
                            {
                                AUTONUMCS = request.AUTONUMCS_OC,
                                ARMAZEM = request.AUTONUM_ARMAZEM,
                                Q = request.QTDE
                            },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        // 3.3) Religa o marcante no novo registro yard. Em parcial, atualiza VOLUMES e
                        // a quantidade na tabela de marcantes temporarios (compat com Associacao de Marcantes).
                        if (!parcial)
                        {
                            con.Execute(
                                @"UPDATE SGIPA..TB_MARCANTES
                                  SET AUTONUM_CS_YARD = @NOVO
                                  WHERE AUTONUM = @MARCANTE",
                                new { NOVO = novoIdYard, MARCANTE = request.MARCANTE },
                                tx,
                                Config.QueryTimeoutInSeconds());
                        }
                        else
                        {
                            con.Execute(
                                @"UPDATE SGIPA..TB_MARCANTES
                                  SET AUTONUM_CS_YARD = @NOVO,
                                      VOLUMES        = @Q
                                  WHERE AUTONUM = @MARCANTE",
                                new
                                {
                                    NOVO = novoIdYard,
                                    Q = request.QTDE,
                                    MARCANTE = request.MARCANTE
                                },
                                tx,
                                Config.QueryTimeoutInSeconds());

                            con.Execute(
                                @"UPDATE SGIPA..tb_temp_marcante
                                  SET QTDE = @Q
                                  WHERE MARCANTE = @MARCANTE_TXT",
                                new
                                {
                                    Q = request.QTDE,
                                    MARCANTE_TXT = request.MARCANTE.ToString("000000000")
                                },
                                tx,
                                Config.QueryTimeoutInSeconds());
                        }

                        // 3.4) Historico do shifting (replica VB6: ARMAZEM=0, YARD='CAM').
                        con.Execute(
                            @"INSERT INTO SGIPA..TB_HIST_SHIFTING_CS
                                  (MARCANTE, ARMAZEM, YARD, DT_MOV, USUARIO)
                              VALUES (@MARCANTE, 0, 'CAM', GETDATE(), @USUARIO)",
                            new { MARCANTE = request.MARCANTE, USUARIO = codUsuario },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        tx.Commit();

                        return new CarregarItemResult
                        {
                            SUCESSO = true,
                            PARCIAL = parcial,
                            MENSAGEM = parcial ? "ITEM PARCIALMENTE CARREGADO" : "ITEM CARREGADO"
                        };
                    }
                    catch (Exception ex)
                    {
                        try { tx.Rollback(); }
                        catch { }
                        return Falha("Erro ao carregar: " + ex.Message);
                    }
                }
            }
        }

        public CancelarItemResult Cancelar(CancelarItemRequest request, int codUsuario)
        {
            if (request == null || request.MARCANTE <= 0)
            {
                return new CancelarItemResult { SUCESSO = false, MENSAGEM = "Marcante nao informado/invalido" };
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();

                MarcanteEstadoDto estado = ObterEstadoMarcante(con, request.MARCANTE, null);
                if (estado == null)
                {
                    return new CancelarItemResult { SUCESSO = false, MENSAGEM = "Marcante nao informado/invalido" };
                }

                if (!string.Equals(estado.YARD, "CAM", StringComparison.OrdinalIgnoreCase))
                {
                    return new CancelarItemResult { SUCESSO = false, MENSAGEM = "Carga ainda nao carregada" };
                }

                using (IDbTransaction tx = con.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        // 1) Zera o registro yard CAM atual.
                        con.Execute(
                            "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE = 0 WHERE AUTONUM = @ID",
                            new { ID = request.AUTONUM_CS_YARD },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        // 2) Insere novo registro yard 'CANCC' com a quantidade cancelada.
                        int novoIdYard = con.QuerySingle<int>(
                            @"INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD
                                  (AUTONUM_CS, ARMAZEM, YARD, ORIGEM, QUANTIDADE, MOTIVO)
                              OUTPUT INSERTED.AUTONUM
                              VALUES (@AUTONUMCS, @ARMAZEM, 'CANCC', 'I', @Q, 8)",
                            new
                            {
                                AUTONUMCS = request.AUTONUMCS_OC,
                                ARMAZEM = request.AUTONUM_ARMAZEM,
                                Q = request.QTDE
                            },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        // 3) Religa o marcante e libera a placa (PLACA_C = NULL).
                        con.Execute(
                            @"UPDATE SGIPA..TB_MARCANTES
                              SET AUTONUM_CS_YARD = @NOVO,
                                  PLACA_C        = NULL
                              WHERE AUTONUM = @MARCANTE",
                            new { NOVO = novoIdYard, MARCANTE = request.MARCANTE },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        // 4) Decisao de migracao: gravar historico tambem no cancelamento (YARD='CANCC').
                        // O legado VB6 nao registra o evento - aqui adicionamos rastreabilidade.
                        con.Execute(
                            @"INSERT INTO SGIPA..TB_HIST_SHIFTING_CS
                                  (MARCANTE, ARMAZEM, YARD, DT_MOV, USUARIO)
                              VALUES (@MARCANTE, 0, 'CANCC', GETDATE(), @USUARIO)",
                            new { MARCANTE = request.MARCANTE, USUARIO = codUsuario },
                            tx,
                            Config.QueryTimeoutInSeconds());

                        tx.Commit();

                        return new CancelarItemResult { SUCESSO = true, MENSAGEM = "ITEM CANCELADO" };
                    }
                    catch (Exception ex)
                    {
                        try { tx.Rollback(); }
                        catch { }
                        return new CancelarItemResult { SUCESSO = false, MENSAGEM = "Erro ao cancelar: " + ex.Message };
                    }
                }
            }
        }

        public DocumentoOcResult ObterDocumento(int loteBl)
        {
            if (loteBl <= 0)
            {
                return new DocumentoOcResult { SUCESSO = false, MENSAGEM = "Escolha uma carga" };
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                int flagLtl = con.QueryFirstOrDefault<int?>(
                    @"SELECT ISNULL(MAX(A.flag_ltl), 0)
                      FROM SGIPA..TB_SOLICITACAO_LTL A
                      INNER JOIN SGIPA..TB_BL B ON A.LOTE = B.AUTONUM
                      WHERE B.AUTONUM = @LOTE",
                    new { LOTE = loteBl },
                    commandTimeout: Config.QueryTimeoutInSeconds()) ?? 0;

                if (flagLtl == 1)
                {
                    return new DocumentoOcResult
                    {
                        SUCESSO = true,
                        LTL = true,
                        LOTE_BL = loteBl,
                        MENSAGEM = "Documentos LTL disponiveis para o lote " + loteBl
                    };
                }

                int? autonumAg = con.QueryFirstOrDefault<int?>(
                    "SELECT MAX(AUTONUM) FROM SGIPA..TB_AG_CS WHERE LOTE = @LOTE",
                    new { LOTE = loteBl },
                    commandTimeout: Config.QueryTimeoutInSeconds());

                if (!autonumAg.HasValue || autonumAg.Value <= 0)
                {
                    return new DocumentoOcResult
                    {
                        SUCESSO = false,
                        MENSAGEM = "Nenhum agendamento de documento encontrado para o lote informado"
                    };
                }

                return new DocumentoOcResult
                {
                    SUCESSO = true,
                    LTL = false,
                    AUTONUM_AG_CS = autonumAg,
                    LOTE_BL = loteBl,
                    MENSAGEM = "Agendamento " + autonumAg.Value + " localizado"
                };
            }
        }

        // -----------------------------------------------------------------
        // Helpers internos
        // -----------------------------------------------------------------

        private static CarregarItemResult ValidarRequestCarregar(CarregarItemRequest request)
        {
            if (request == null)
            {
                return Falha("Requisicao invalida");
            }

            if (request.MARCANTE <= 0)
            {
                return Falha("Marcante nao informado/invalido");
            }

            if (request.AUTONUMCS_OC <= 0 || request.ORDEM <= 0)
            {
                return Falha("Ordem de carregamento nao selecionada");
            }

            if (request.QTDE <= 0)
            {
                return Falha("Quantidade indisponivel");
            }

            if (request.AUTONUM_CS_YARD <= 0)
            {
                return Falha("Marcante nao informado/invalido");
            }

            if (string.IsNullOrWhiteSpace(request.PLACA_C))
            {
                return Falha("Veiculo nao selecionado");
            }

            return null;
        }

        private static CarregarItemResult Falha(string mensagem)
        {
            return new CarregarItemResult { SUCESSO = false, MENSAGEM = mensagem };
        }

        private static MarcanteEstadoDto ObterEstadoMarcante(SqlConnection con, int marcante, IDbTransaction tx)
        {
            const string sql = @"
                SELECT ISNULL(M.AUTONUM_CARGA, 0)  AS AUTONUMCS,
                       ISNULL(M.VOLUMES, 0)        AS VOLUMES,
                       ISNULL(M.AUTONUM_CS_YARD,0) AS AUTONUM_CS_YARD,
                       ISNULL(Y.YARD, '')          AS YARD
                FROM   SGIPA..TB_MARCANTES M
                LEFT   JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD = Y.AUTONUM
                WHERE  M.AUTONUM = @MARCANTE";

            return con.QueryFirstOrDefault<MarcanteEstadoDto>(
                sql,
                new { MARCANTE = marcante },
                tx,
                Config.QueryTimeoutInSeconds());
        }

        private static string LimitarPlaca(string placa)
        {
            if (string.IsNullOrEmpty(placa))
            {
                return null;
            }

            string trimmed = placa.Trim();
            return trimmed.Length > 8 ? trimmed.Substring(0, 8) : trimmed;
        }

        private class MarcanteEstadoDto
        {
            public int AUTONUMCS { get; set; }
            public int VOLUMES { get; set; }
            public int AUTONUM_CS_YARD { get; set; }
            public string YARD { get; set; }
        }
    }
}
