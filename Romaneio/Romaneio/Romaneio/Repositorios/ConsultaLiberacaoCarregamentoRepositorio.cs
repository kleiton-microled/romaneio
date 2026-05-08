using Dapper;
using Romaneio.Helpers;
using Romaneio.Interfaces;
using Romaneio.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;

namespace Romaneio.Repositorios
{
    public class ConsultaLiberacaoCarregamentoRepositorio : IConsultaLiberacaoCarregamentoRepositorio
    {
        public List<TipoVeiculoDto> ObterTiposVeiculo()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                const string sql = @"SELECT AUTONUM, DESCR AS DISPLAY
                                     FROM SGIPA..TB_TIPOS_CAMINHAO
                                     ORDER BY DESCR";
                return con.Query<TipoVeiculoDto>(sql, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }
        }

        public ConsultaLiberacaoResultadoDto Filtrar(ConsultaLiberacaoFiltroRequest request, int patioUsuario)
        {
            request = request ?? new ConsultaLiberacaoFiltroRequest();
            var protocoloNumerico = SomenteNumeros(request.PROTOCOLO);
            var placa = (request.PLACA ?? string.Empty).Trim().ToUpperInvariant();
            var cntr = (request.CNTR ?? string.Empty).Trim().ToUpperInvariant();
            var bl = (request.BL ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(protocoloNumerico) && string.IsNullOrWhiteSpace(placa) && string.IsNullOrWhiteSpace(bl))
            {
                return Falha("Informe o BL, placa ou protocolo");
            }

            if (string.IsNullOrWhiteSpace(protocoloNumerico))
            {
                if (string.IsNullOrWhiteSpace(placa))
                {
                    return Falha("Informe a placa");
                }

                if (string.IsNullOrWhiteSpace(cntr) && string.IsNullOrWhiteSpace(bl))
                {
                    return Falha("Informe o conteiner ou o BL");
                }
            }

            var protocoloFormatado = FormatarProtocolo(protocoloNumerico);

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var dados = ConsultarAgendaPrincipal(con, protocoloFormatado, placa, bl);
                if (dados == null)
                {
                    dados = ConsultarFallbackCntr(con, protocoloFormatado, placa, cntr, bl);
                }

                if (dados == null)
                {
                    return Falha("Registro de Saida nao localizado favor entrar no sistema Gerencial e efetuar o registro de saida da carga");
                }

                if (!PatioCompativel(patioUsuario, dados.PATIO))
                {
                    return Falha("Patio divergente do usuario.");
                }

                var result = new ConsultaLiberacaoResultadoDto
                {
                    SUCESSO = true,
                    PROTOCOLO = string.IsNullOrWhiteSpace(dados.PROTOCOLO) ? protocoloFormatado : dados.PROTOCOLO,
                    PERIODO = dados.PERIODO,
                    LOTE = dados.LOTE.ToString(),
                    AUTONUM_LOTE = dados.LOTE,
                    NUM_DOCUMENTO = dados.NUM_DOCUMENTO,
                    TRANSPORTADORA = dados.TRANSPORTADORA,
                    MOTORISTA = dados.NOME_MOTORISTA,
                    CPF_MOTORISTA = dados.CPFMOT,
                    PLACAS = $"{(dados.PLACA_CAVALO ?? string.Empty).Trim()} - {(dados.PLACA_CARRETA ?? string.Empty).Trim()}".Trim(' ', '-'),
                    ORDEM_CARREGAMENTO = dados.ORDEM,
                    AUTONUM_CNTR = dados.AUTONUM_CNTR,
                    AUTONUM_CS = dados.AUTONUM_CS,
                    TIPO_CAMINHAO = dados.ID_TIPO_CAMINHAO,
                    STATUS_SISCARGA = "NAO ANALISADO",
                    STATUS_SEFAZ = "NAO ANALISADO"
                };

                result.CARGAS = CarregarCargas(con, dados.ORDEM, dados.AUTONUM_CS);
                result.PROTOCOLOS_SIMULTANEOS = CarregarProtocolosSimultaneosInterno(con, result.PROTOCOLO);
                result.LOCALIZACAO = CarregarLocalizacao(con, dados.AUTONUM_CS);

                result.BLOQUEIO_BL = ConsultaBloqueioBl(con, dados.LOTE);
                result.BLOQUEIO_CNTR = dados.AUTONUM_CNTR > 0 && ConsultaBloqueioCntr(con, dados.AUTONUM_CNTR.Value);
                result.GLME_PENDENTE = ConsultaGlmePendente(con, dados.ORDEM);

                var flagPatioLiberacao = ConsultaFlagLiberacaoPatio(con, dados.LOTE);
                result.PODE_REGISTRAR = flagPatioLiberacao && !result.BLOQUEIO_BL && !result.BLOQUEIO_CNTR && !result.GLME_PENDENTE;
                result.MENSAGEM = result.PODE_REGISTRAR ? string.Empty : "Carga localizada com restricoes para registro";
                return result;
            }
        }

        public List<string> CarregarProtocolosSimultaneos(string protocolo)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                return CarregarProtocolosSimultaneosInterno(con, protocolo);
            }
        }

        public RegistrarLiberacaoResultDto Registrar(RegistrarLiberacaoRequest request, int usuarioId, int patioUsuario)
        {
            if (request == null || request.ORDEM_CARREGAMENTO <= 0 || request.AUTONUM_LOTE <= 0)
            {
                return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Registro de Saida nao localizado..." };
            }

            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                con.Open();
                using (var transaction = con.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        var patioCarga = con.QuerySingleOrDefault<int>(
                            "SELECT ISNULL(PATIO,0) FROM SGIPA..TB_BL WHERE AUTONUM = @LOTE",
                            new { LOTE = request.AUTONUM_LOTE },
                            transaction,
                            Config.QueryTimeoutInSeconds());

                        if (!PatioCompativel(patioUsuario, patioCarga))
                        {
                            transaction.Rollback();
                            return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Patio divergente do usuario." };
                        }

                        if (ConsultaBloqueioBl(con, request.AUTONUM_LOTE, transaction))
                        {
                            transaction.Rollback();
                            return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Existe Bloqueio de BL" };
                        }

                        if (request.AUTONUM_CNTR.HasValue && request.AUTONUM_CNTR.Value > 0 && ConsultaBloqueioCntr(con, request.AUTONUM_CNTR.Value, transaction))
                        {
                            transaction.Rollback();
                            return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Existe Bloqueio de Conteiner" };
                        }

                        if (ConsultaGlmePendente(con, request.ORDEM_CARREGAMENTO, transaction))
                        {
                            transaction.Rollback();
                            return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Atencao, Existe Carga sem Liberacao GLME ." };
                        }

                        var flagPatio = ConsultaFlagLiberacaoPatio(con, request.AUTONUM_LOTE, transaction);
                        if (!flagPatio)
                        {
                            transaction.Rollback();
                            return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Patio parametrizado para nao realizar consultas!" };
                        }

                        if (!string.IsNullOrWhiteSpace(request.PLACA_CAVALO) && !string.IsNullOrWhiteSpace(request.PLACA_CARRETA))
                        {
                            con.Execute(
                                @"UPDATE SGIPA..TB_AG_VEICULOS
                                  SET ID_TIPO_CAMINHAO = @TIPO
                                  WHERE PLACA_CAVALO = @PLACA_CAVALO
                                    AND PLACA_CARRETA = @PLACA_CARRETA",
                                new
                                {
                                    TIPO = request.TIPO_CAMINHAO,
                                    PLACA_CAVALO = request.PLACA_CAVALO,
                                    PLACA_CARRETA = request.PLACA_CARRETA
                                },
                                transaction,
                                Config.QueryTimeoutInSeconds());
                        }

                        var atualizados = con.Execute(
                            @"UPDATE SGIPA..TB_ORDEM_CARREGAMENTO
                              SET USUARIO_REGISTRO_CAM = @USUARIO,
                                  DATA_ORDEM = ISNULL(DATA_ORDEM, GETDATE())
                              WHERE AUTONUM = @ORDEM",
                            new
                            {
                                USUARIO = usuarioId,
                                ORDEM = request.ORDEM_CARREGAMENTO
                            },
                            transaction,
                            Config.QueryTimeoutInSeconds());

                        if (atualizados <= 0)
                        {
                            transaction.Rollback();
                            return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Nao foi possivel registrar a saida da carga!" };
                        }

                        transaction.Commit();
                        return new RegistrarLiberacaoResultDto { SUCESSO = true, MENSAGEM = "Registrado com sucesso" };
                    }
                    catch
                    {
                        transaction.Rollback();
                        return new RegistrarLiberacaoResultDto { SUCESSO = false, MENSAGEM = "Nao foi possivel registrar a saida da carga!" };
                    }
                }
            }
        }

        private static ConsultaLiberacaoResultadoDto Falha(string mensagem)
        {
            return new ConsultaLiberacaoResultadoDto
            {
                SUCESSO = false,
                MENSAGEM = mensagem
            };
        }

        private static string SomenteNumeros(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return string.Empty;
            }

            return Regex.Replace(valor, @"\D", string.Empty);
        }

        private static string FormatarProtocolo(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return string.Empty;
            }

            var digitos = SomenteNumeros(valor).PadLeft(10, '0');
            return $"{digitos.Substring(0, 6)}/{digitos.Substring(6, 4)}";
        }

        private static bool PatioCompativel(int patioUsuario, int patioCarga)
        {
            if (patioUsuario == 1 || patioUsuario == 7)
            {
                return patioCarga == 1 || patioCarga == 7;
            }

            return patioUsuario == patioCarga;
        }

        private static List<CargaListaItemDto> CarregarCargas(SqlConnection con, int ordem, int? autonumCs)
        {
            if (ordem > 0)
            {
                const string sqlOrdem = @"SELECT DISTINCT
                                              CASE
                                                  WHEN A.AUTONUM IS NOT NULL
                                                      THEN 'Carga Solta - ' + CONVERT(VARCHAR, A.ITEM) + ' - ' + ISNULL(E.DESCR,'')
                                                  ELSE C.ID_CONTEINER
                                              END AS CARGA
                                          FROM SGIPA..TB_ORDEM_CARREGAMENTO O
                                          LEFT JOIN SGIPA..TB_REGISTRO_SAIDA_CS RCS ON RCS.ORDEM_CARREG = O.AUTONUM
                                          LEFT JOIN SGIPA..TB_CARGA_SOLTA A ON A.AUTONUM = RCS.CS
                                          LEFT JOIN SGIPA..DTE_TB_EMBALAGENS E ON E.CODE = A.EMBALAGEM
                                          LEFT JOIN SGIPA..TB_REGISTRO_SAIDA_CNTR RCN ON RCN.ORDEM_CARREG = O.AUTONUM
                                          LEFT JOIN SGIPA..TB_CNTR_BL C ON C.AUTONUM = RCN.CNTR
                                          WHERE O.AUTONUM = @ORDEM";

                return con.Query<CargaListaItemDto>(sqlOrdem, new { ORDEM = ordem }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }

            if (autonumCs.HasValue && autonumCs.Value > 0)
            {
                const string sqlCs = @"SELECT 'Carga Solta - ' + CONVERT(VARCHAR,A.ITEM) + ' - ' + ISNULL(E.DESCR,'') AS CARGA
                                       FROM SGIPA..TB_CARGA_SOLTA A
                                       LEFT JOIN SGIPA..DTE_TB_EMBALAGENS E ON E.CODE = A.EMBALAGEM
                                       WHERE A.AUTONUM = @AUTONUM_CS";

                return con.Query<CargaListaItemDto>(sqlCs, new { AUTONUM_CS = autonumCs.Value }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
            }

            return new List<CargaListaItemDto>();
        }

        private static List<string> CarregarProtocolosSimultaneosInterno(SqlConnection con, string protocolo)
        {
            if (string.IsNullOrWhiteSpace(protocolo))
            {
                return new List<string>();
            }

            const string sql = @"SELECT DISTINCT B.PROTOCOLO
                                 FROM SGIPA..VW_AGENDA_CS A
                                 INNER JOIN SGIPA..VW_AGENDA_CS B
                                         ON A.PLACA_CAVALO = B.PLACA_CAVALO
                                        AND A.PERIODO = B.PERIODO
                                 WHERE A.PROTOCOLO = @PROTOCOLO
                                   AND B.PROTOCOLO <> @PROTOCOLO
                                 ORDER BY B.PROTOCOLO";

            return con.Query<string>(sql, new { PROTOCOLO = protocolo }, commandTimeout: Config.QueryTimeoutInSeconds()).ToList();
        }

        private static string CarregarLocalizacao(SqlConnection con, int? autonumCs)
        {
            if (!autonumCs.HasValue || autonumCs.Value <= 0)
            {
                return string.Empty;
            }

            const string sql = @"SELECT TOP 1 (DESCR_ARMAZEM + ' ' + POSICAO)
                                 FROM OPERADOR..VW_INVENT_ARMAZEM_ITEM
                                 WHERE AUTONUMCS = @AUTONUMCS
                                 ORDER BY QTDE DESC";

            return con.QueryFirstOrDefault<string>(sql, new { AUTONUMCS = autonumCs.Value }, commandTimeout: Config.QueryTimeoutInSeconds()) ?? string.Empty;
        }

        private static bool ConsultaBloqueioBl(SqlConnection con, int lote, IDbTransaction transaction = null)
        {
            const string sql = @"SELECT ISNULL(MAX(ISNULL(FLAG_BLOQUEIO,0) + ISNULL(FLAG_BLOQUEIO_MANUAL,0)),0)
                                 FROM SGIPA..TB_BL
                                 WHERE AUTONUM = @LOTE";
            return con.QuerySingleOrDefault<int>(sql, new { LOTE = lote }, transaction, Config.QueryTimeoutInSeconds()) > 0;
        }

        private static bool ConsultaBloqueioCntr(SqlConnection con, int autonumCntr, IDbTransaction transaction = null)
        {
            const string sql = @"SELECT ISNULL(FLAG_BLOQUEIO,0)
                                 FROM SGIPA..TB_CNTR_BL
                                 WHERE AUTONUM = @AUTONUM_CNTR";
            return con.QuerySingleOrDefault<int>(sql, new { AUTONUM_CNTR = autonumCntr }, transaction, Config.QueryTimeoutInSeconds()) > 0;
        }

        private static bool ConsultaGlmePendente(SqlConnection con, int ordem, IDbTransaction transaction = null)
        {
            if (ordem <= 0)
            {
                return false;
            }

            const string sqlCs = @"SELECT COUNT(1)
                                   FROM SGIPA..TB_REGISTRO_SAIDA_CS
                                   WHERE ISNULL(FLAG_GLME,0) = 0
                                     AND ORDEM_CARREG = @ORDEM
                                     AND CS IN (
                                         SELECT AUTONUM
                                         FROM SGIPA..TB_CARGA_SOLTA
                                         WHERE BL IN (
                                             SELECT BL
                                             FROM SGIPA..TB_ETAPAS_CARGA
                                             WHERE ISNULL(FLAG_GLME,0) = 1
                                         )
                                     )";

            const string sqlCntr = @"SELECT COUNT(1)
                                     FROM SGIPA..TB_REGISTRO_SAIDA_CNTR
                                     WHERE ISNULL(FLAG_GLME,0) = 0
                                       AND ORDEM_CARREG = @ORDEM
                                       AND CNTR IN (
                                           SELECT CNTR
                                           FROM SGIPA..TB_AMR_CNTR_BL
                                           WHERE BL IN (
                                               SELECT BL
                                               FROM SGIPA..TB_ETAPAS_CARGA
                                               WHERE ISNULL(FLAG_GLME,0) = 1
                                           )
                                       )";

            var pendCs = con.QuerySingleOrDefault<int>(sqlCs, new { ORDEM = ordem }, transaction, Config.QueryTimeoutInSeconds());
            var pendCntr = con.QuerySingleOrDefault<int>(sqlCntr, new { ORDEM = ordem }, transaction, Config.QueryTimeoutInSeconds());
            return pendCs > 0 || pendCntr > 0;
        }

        private static bool ConsultaFlagLiberacaoPatio(SqlConnection con, int lote, IDbTransaction transaction = null)
        {
            const string sql = @"SELECT ISNULL(P.FLAG_CONSULTA_LIBERACAO_COLETOR,0)
                                 FROM SGIPA..TB_BL BL
                                 LEFT JOIN OPERADOR..TB_PATIOS P ON BL.PATIO = P.AUTONUM
                                 WHERE BL.AUTONUM = @LOTE";

            return con.QuerySingleOrDefault<int>(sql, new { LOTE = lote }, transaction, Config.QueryTimeoutInSeconds()) > 0;
        }

        private static AgendaConsultaDto ConsultarAgendaPrincipal(SqlConnection con, string protocolo, string placa, string bl)
        {
            var sql = @"SELECT TOP 1
                            AG.PROTOCOLO,
                            AG.PERIODO,
                            AG.LOTE,
                            AG.NUM_DOCUMENTO,
                            C.RAZAO AS TRANSPORTADORA,
                            AG.NOME_MOTORISTA,
                            AG.PLACA_CAVALO,
                            AG.PLACA_CARRETA,
                            AG.AUTONUM_CS,
                            0 AS AUTONUM_CNTR,
                            ISNULL(AG.ID_TIPO_CAMINHAO,0) AS ID_TIPO_CAMINHAO,
                            BL.PATIO,
                            B.CPF AS CPFMOT,
                            ISNULL((
                                SELECT MIN(ORDEM_CARREG)
                                FROM SGIPA..TB_REGISTRO_SAIDA_CS
                                WHERE CS = AG.AUTONUM_CS
                            ), 0) AS ORDEM
                        FROM SGIPA..VW_AGENDA_CS AG
                        INNER JOIN OPERADOR..TB_CAD_TRANSPORTADORAS C ON AG.COD_TRANSPORTADORA = C.AUTONUM
                        INNER JOIN SGIPA..TB_BL BL ON AG.LOTE = BL.AUTONUM
                        LEFT JOIN OPERADOR..TB_MOTORISTAS B ON AG.CNH = B.CNH
                        WHERE 1=1";

            var parametros = new DynamicParameters();
            if (!string.IsNullOrWhiteSpace(protocolo))
            {
                sql += " AND AG.PROTOCOLO = @PROTOCOLO";
                parametros.Add("PROTOCOLO", protocolo);
            }

            if (!string.IsNullOrWhiteSpace(placa))
            {
                sql += " AND AG.PLACA_CAVALO = @PLACA";
                parametros.Add("PLACA", placa);
            }

            if (!string.IsNullOrWhiteSpace(bl))
            {
                sql += " AND BL.NUM_DOCUMENTO LIKE '%' + @BL + '%'";
                parametros.Add("BL", bl);
            }

            sql += " ORDER BY AG.PERIODO DESC";
            return con.QueryFirstOrDefault<AgendaConsultaDto>(sql, parametros, commandTimeout: Config.QueryTimeoutInSeconds());
        }

        private static AgendaConsultaDto ConsultarFallbackCntr(SqlConnection con, string protocolo, string placa, string cntr, string bl)
        {
            var sql = @"SELECT TOP 1
                            CONVERT(VARCHAR, C.NUM_PROTOCOLO) + '/' + CONVERT(VARCHAR, C.ANO_PROTOCOLO) AS PROTOCOLO,
                            CONVERT(VARCHAR, GD.PERIODO_INICIAL, 103) + ' ' + CONVERT(VARCHAR(5),GD.PERIODO_INICIAL,108) + ' - ' +
                            CONVERT(VARCHAR, GD.PERIODO_FINAL, 103) + ' ' + CONVERT(VARCHAR(5),GD.PERIODO_FINAL,108) AS PERIODO,
                            BL.AUTONUM AS LOTE,
                            BL.NUM_DOCUMENTO,
                            T.RAZAO AS TRANSPORTADORA,
                            B.NOME AS NOME_MOTORISTA,
                            OC.PLACA_C AS PLACA_CAVALO,
                            OC.PLACA_CARRETA,
                            0 AS AUTONUM_CS,
                            C.AUTONUM AS AUTONUM_CNTR,
                            ISNULL(AG.ID_TIPO_CAMINHAO,0) AS ID_TIPO_CAMINHAO,
                            BL.PATIO,
                            B.CPF AS CPFMOT,
                            ISNULL(OC.AUTONUM,0) AS ORDEM
                        FROM SGIPA..TB_CNTR_BL C
                        INNER JOIN SGIPA..TB_AMR_CNTR_BL AMR ON C.AUTONUM = AMR.CNTR
                        INNER JOIN SGIPA..TB_BL BL ON AMR.BL = BL.AUTONUM
                        LEFT JOIN OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA = GD.AUTONUM_GD_RESERVA
                        INNER JOIN SGIPA..TB_REGISTRO_SAIDA_CNTR SCC ON C.AUTONUM = SCC.CNTR
                        INNER JOIN SGIPA..TB_ORDEM_CARREGAMENTO OC ON SCC.ORDEM_CARREG = OC.AUTONUM
                        INNER JOIN OPERADOR..TB_MOTORISTAS B ON OC.ID_MOTORISTA = B.AUTONUM
                        INNER JOIN OPERADOR..TB_CAD_TRANSPORTADORAS T ON OC.TRANSPORTADORA = T.AUTONUM
                        LEFT JOIN SGIPA..TB_AG_VEICULOS AG ON C.AUTONUM_VEICULO = AG.AUTONUM
                        WHERE BL.FLAG_ATIVO = 1";

            var parametros = new DynamicParameters();
            if (!string.IsNullOrWhiteSpace(protocolo))
            {
                var partes = protocolo.Split('/');
                if (partes.Length == 2)
                {
                    sql += " AND C.NUM_PROTOCOLO = @NUM_PROTOCOLO AND C.ANO_PROTOCOLO = @ANO_PROTOCOLO";
                    parametros.Add("NUM_PROTOCOLO", Convert.ToInt32(partes[0]));
                    parametros.Add("ANO_PROTOCOLO", Convert.ToInt32(partes[1]));
                }
            }

            if (!string.IsNullOrWhiteSpace(placa))
            {
                sql += " AND OC.PLACA_C = @PLACA";
                parametros.Add("PLACA", placa);
            }

            if (!string.IsNullOrWhiteSpace(cntr))
            {
                sql += " AND SUBSTRING(C.ID_CONTEINER,5,8) = @CNTR";
                parametros.Add("CNTR", cntr.Replace("-", string.Empty));
            }

            if (!string.IsNullOrWhiteSpace(bl))
            {
                sql += " AND BL.NUM_DOCUMENTO LIKE '%' + @BL + '%'";
                parametros.Add("BL", bl);
            }

            sql += " ORDER BY GD.PERIODO_INICIAL DESC";
            return con.QueryFirstOrDefault<AgendaConsultaDto>(sql, parametros, commandTimeout: Config.QueryTimeoutInSeconds());
        }

        private class AgendaConsultaDto
        {
            public string PROTOCOLO { get; set; }
            public string PERIODO { get; set; }
            public int LOTE { get; set; }
            public string NUM_DOCUMENTO { get; set; }
            public string TRANSPORTADORA { get; set; }
            public string NOME_MOTORISTA { get; set; }
            public string PLACA_CAVALO { get; set; }
            public string PLACA_CARRETA { get; set; }
            public int AUTONUM_CS { get; set; }
            public int? AUTONUM_CNTR { get; set; }
            public int ID_TIPO_CAMINHAO { get; set; }
            public int PATIO { get; set; }
            public string CPFMOT { get; set; }
            public int ORDEM { get; set; }
        }
    }
}
