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
    /// <summary>
    /// Porta a logica de listagem de documentos do projeto WinForms ConsultaDocColetor
    /// (FrmDocumentosAgendamento.Consultar / ConsultarDocumentosLTL) para SQL Server.
    /// A obtencao de arquivo via servico Bandeirantes.AnexoDiretorio nao esta no Romaneio:
    /// quando DOCUMENTO na TB_AG_DOCUMENTOS for nulo, o download direto nao estara disponivel.
    /// </summary>
    public class ConsultaDocumentosAgendamentoRepositorio : IConsultaDocumentosAgendamentoRepositorio
    {
        private static string FmtDtInclusao(string alias)
        {
            return "CONVERT(VARCHAR(10), " + alias + ".DT_INCLUSAO, 103) + ' ' + LEFT(CONVERT(VARCHAR(8), " + alias + ".DT_INCLUSAO, 108), 5)";
        }

        public int? ResolverIdSolicitacaoLtlPorLote(int lote)
        {
            if (lote <= 0)
            {
                return null;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                return con.QueryFirstOrDefault<int?>(
                    @"SELECT TOP 1 A.ID
                      FROM SGIPA..TB_SOLICITACAO_LTL A
                      WHERE A.LOTE = @LOTE
                      ORDER BY A.ID DESC",
                    new { LOTE = lote },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public List<DocumentoAgendamentoListaDto> ListarDocumentosCargaSolta(int autonumAgendamento, int lote)
        {
            if (autonumAgendamento <= 0 || lote <= 0)
            {
                return new List<DocumentoAgendamentoListaDto>();
            }

            string dt = FmtDtInclusao("A");

            var sql = new StringBuilder();
            sql.AppendLine("SELECT DISTINCT DESCRICAO, LOTE, NOME_IMG, DT_INCLUSAO, AUTONUM_AV_IMAGEM, TIPO_DOCUMENTO FROM (");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_CS AG ON A.AUTONUM_AGENDAMENTO = AG.AUTONUM
                WHERE AG.LOTE = @LOTE AND A.TIPO = 'CS' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine("  UNION ALL");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_MOTORISTAS C ON A.AUTONUM = C.DOC_MOTORISTA
                LEFT JOIN SGIPA..TB_AG_CS AG ON A.AUTONUM_AGENDAMENTO = AG.AUTONUM
                WHERE A.TIPO = 'CS' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine("  UNION ALL");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_MOTORISTAS C ON A.AUTONUM = C.DOC_MOTORISTA
                LEFT JOIN SGIPA..TB_CNTR_BL D ON A.AUTONUM_AGENDAMENTO = D.AUTONUM
                LEFT JOIN (SELECT MAX(BL) AS LOTE, CNTR FROM SGIPA..TB_AMR_CNTR_BL GROUP BY CNTR) AG ON D.AUTONUM = AG.CNTR
                WHERE D.AUTONUM = @AG AND A.TIPO = 'CN' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine("  UNION ALL");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_VEICULOS C ON A.AUTONUM = C.DOC_CAVALO
                LEFT JOIN SGIPA..TB_AG_CS AG ON A.AUTONUM_AGENDAMENTO = AG.AUTONUM
                WHERE A.TIPO = 'CS' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine("  UNION ALL");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_VEICULOS C ON A.AUTONUM = C.DOC_CAVALO
                LEFT JOIN SGIPA..TB_CNTR_BL D ON A.AUTONUM_AGENDAMENTO = D.AUTONUM
                LEFT JOIN (SELECT MAX(BL) AS LOTE, CNTR FROM SGIPA..TB_AMR_CNTR_BL GROUP BY CNTR) AG ON D.AUTONUM = AG.CNTR
                WHERE A.TIPO = 'CNTR' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine("  UNION ALL");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_VEICULOS C ON A.AUTONUM = C.DOC_CARRETA
                LEFT JOIN SGIPA..TB_AG_CS AG ON A.AUTONUM_AGENDAMENTO = AG.AUTONUM
                WHERE A.TIPO = 'CS' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine("  UNION ALL");
            sql.AppendLine("  SELECT DISTINCT B.DESCRICAO, AG.LOTE, A.NOME_IMG, " + dt + @" AS DT_INCLUSAO, A.AUTONUM AS AUTONUM_AV_IMAGEM, B.AUTONUM AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_AG_VEICULOS C ON A.AUTONUM = C.DOC_CARRETA
                LEFT JOIN SGIPA..TB_CNTR_BL D ON A.AUTONUM_AGENDAMENTO = D.AUTONUM
                LEFT JOIN (SELECT MAX(BL) AS LOTE, CNTR FROM SGIPA..TB_AMR_CNTR_BL GROUP BY CNTR) AG ON D.AUTONUM = AG.CNTR
                WHERE A.TIPO = 'CNTR' AND A.AUTONUM_AGENDAMENTO = @AG");
            sql.AppendLine(") X ORDER BY DT_INCLUSAO DESC, DESCRICAO");

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                return con.Query<DocumentoAgendamentoListaDto>(
                    sql.ToString(),
                    new { AG = autonumAgendamento, LOTE = lote },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public List<DocumentoAgendamentoListaDto> ListarDocumentosLtl(int idSolicitacaoLtl, int lote)
        {
            if (idSolicitacaoLtl <= 0 || lote <= 0)
            {
                return new List<DocumentoAgendamentoListaDto>();
            }

            const string sql = @"SELECT
                    B.DESCRICAO,
                    AG.LOTE,
                    A.NOME_IMG,
                    CONVERT(VARCHAR(10), A.DT_INCLUSAO, 103) + ' ' + LEFT(CONVERT(VARCHAR(8), A.DT_INCLUSAO, 108), 5) AS DT_INCLUSAO,
                    A.AUTONUM AS AUTONUM_AV_IMAGEM,
                    CAST(NULL AS INT) AS TIPO_DOCUMENTO
                FROM SGIPA..TB_AG_DOCUMENTOS A
                INNER JOIN SGIPA..TB_AGENDAMENTO_DOC B ON A.AUTONUM_TIPO_DOCUMENTO = B.AUTONUM
                INNER JOIN SGIPA..TB_SOLICITACAO_LTL AG ON A.AUTONUM_AGENDAMENTO = AG.ID
                WHERE AG.LOTE = @LOTE
                  AND A.TIPO = 'LTL'
                  AND A.AUTONUM_AGENDAMENTO = @ID";

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                return con.Query<DocumentoAgendamentoListaDto>(
                    sql,
                    new { ID = idSolicitacaoLtl, LOTE = lote },
                    commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public DocumentoAgendamentoArquivoMetaDto ObterMetaArquivo(int autonumDocumento, int autonumAgendamento, bool ltl, int lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                if (ltl)
                {
                    return con.QueryFirstOrDefault<DocumentoAgendamentoArquivoMetaDto>(
                        @"SELECT UPPER(ISNULL(A.EXTENSAO, '')) AS EXTENSAO,
                                 CASE WHEN A.DOCUMENTO IS NULL THEN 0 ELSE 1 END AS TEM_BLOB
                          FROM SGIPA..TB_AG_DOCUMENTOS A
                          INNER JOIN SGIPA..TB_SOLICITACAO_LTL L ON A.AUTONUM_AGENDAMENTO = L.ID
                          WHERE A.AUTONUM = @DOC
                            AND A.TIPO = 'LTL'
                            AND L.LOTE = @LOTE
                            AND L.ID = @ID",
                        new { DOC = autonumDocumento, LOTE = lote, ID = autonumAgendamento },
                        commandTimeout: Config.QueryTimeoutInSeconds());
                }

                return con.QueryFirstOrDefault<DocumentoAgendamentoArquivoMetaDto>(
                    @"SELECT UPPER(ISNULL(A.EXTENSAO, '')) AS EXTENSAO,
                             CASE WHEN A.DOCUMENTO IS NULL THEN 0 ELSE 1 END AS TEM_BLOB
                      FROM SGIPA..TB_AG_DOCUMENTOS A
                      WHERE A.AUTONUM = @DOC
                        AND A.AUTONUM_AGENDAMENTO = @AG",
                    new { DOC = autonumDocumento, AG = autonumAgendamento },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public byte[] ObterBlobDocumento(int autonumDocumento)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                return con.QueryFirstOrDefault<byte[]>(
                    "SELECT DOCUMENTO FROM SGIPA..TB_AG_DOCUMENTOS WHERE AUTONUM = @DOC AND DOCUMENTO IS NOT NULL",
                    new { DOC = autonumDocumento },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public void InserirBloqueioDocumento(int usuarioId, int autonumAgendamento)
        {
            if (usuarioId <= 0 || autonumAgendamento <= 0)
            {
                return;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Execute(
                    @"INSERT INTO OPERADOR..TB_ADMPATIO_BLOQUEIO_DOC (FLAG_BLOQUEIO, USUARIO, AUTONUM_AGENDAMENTO, DATA_BLOQUEIO)
                      VALUES (1, @USUARIO, @AG, GETDATE())",
                    new { USUARIO = usuarioId, AG = autonumAgendamento },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }

        public void RemoverBloqueioDocumento(int usuarioId, int autonumAgendamento)
        {
            if (usuarioId <= 0 || autonumAgendamento <= 0)
            {
                return;
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Execute(
                    @"DELETE FROM OPERADOR..TB_ADMPATIO_BLOQUEIO_DOC
                      WHERE USUARIO = @USUARIO AND AUTONUM_AGENDAMENTO = @AG",
                    new { USUARIO = usuarioId, AG = autonumAgendamento },
                    commandTimeout: Config.QueryTimeoutInSeconds());
            }
        }
    }
}
