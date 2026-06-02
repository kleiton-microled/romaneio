using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Romaneio.Helpers;
using Romaneio.Classes;
using Dapper;
using System.Data.SqlClient;
using System.Text;
using Romaneio.Interfaces;
using Romaneio.Models;
using System.Data;

namespace Romaneio.Repositorios
{
    public class RomaneioRepositorio : IRomaneioRepositorio
    {

        public IEnumerable<Conteiners> ConsultarConteiners(int PATIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();

                sb.AppendLine("SELECT DISTINCT AUTONUM_CNTR, ID_CONTEINER, DESCR_TIPO, TAMANHO, DT_ENTRADA, DT_AGENDAMENTO, ");
                sb.AppendLine("LACRE_ORIGEM, LACRE_IPA, VIAGEM, BOL_DESUNIT, DDA, SOLICITADO, DESCRICAO_SIGVIG, POSICAO, ");
                sb.AppendLine("MIN(ABERTURA) AS ABERTURA, MAX(FECHAMENTO) AS FECHAMENTO ");
                sb.AppendLine("FROM VW_ROMANEIO_CONSULTA_LOTE_CNTR WHERE PATIO = " + PATIO + " ");
                sb.AppendLine("GROUP BY AUTONUM_CNTR, ID_CONTEINER, DESCR_TIPO, TAMANHO, DT_ENTRADA, DT_AGENDAMENTO,  LACRE_ORIGEM, LACRE_IPA, VIAGEM, ");
                sb.AppendLine("BOL_DESUNIT, DDA, SOLICITADO, DESCRICAO_SIGVIG, POSICAO ");
                sb.AppendLine("ORDER BY DT_ENTRADA DESC ");

                var query = con.Query<Conteiners>(sb.ToString()).AsEnumerable();

                return query;

            }
        }

        public IEnumerable<LotesDTO> ConsultarLotesConteiners(string cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.AppendLine("SELECT DISTINCT  AUTONUM_CNTR, LOTE, BL, ITEM, MARCA, QUANTIDADE, NCM, MERCADORIA, LOCALIZACAO, IMPORTADOR, SHIPPER   ");
                sb.AppendLine("FROM VW_ROMANEIO_CONSULTA_LOTE_CNTR ");
                sb.AppendLine("WHERE AUTONUM_CNTR = " + cntr);
                sb.AppendLine(" ORDER BY LOTE DESC ");
                var query = con.Query<LotesDTO>(sb.ToString()).AsEnumerable();

                return query;

            }
        }
        public IEnumerable<Lotes> ListaLotes(int cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT LOTE AS AUTONUM, LOTE FROM VW_ROMANEIO_CONSULTA_LOTE_CNTR WHERE AUTONUM_CNTR = " + cntr + " ORDER BY LOTE ");

                var query = con.Query<Lotes>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public LotesDTO ObterDadosLote(int cntr = 0, int lote = 0, int item = 0)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Config.StringConexao()))
                {
                    var parametros = new DynamicParameters();
                    parametros.Add("cntr", cntr, direction: ParameterDirection.Input);
                    parametros.Add("lote", lote, direction: ParameterDirection.Input);
                    parametros.Add("item", item, direction: ParameterDirection.Input);

                    if (item == 0)
                    {
                        return con.Query<LotesDTO>(@"SELECT AUTONUM_CNTR, ID_CONTEINER, LOTE, FLAG_HUB FROM VW_ROMANEIO_CONSULTA_LOTE_CNTR WHERE LOTE = @lote AND AUTONUM_CNTR = @cntr ORDER BY LOTE DESC", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    }
                    else
                    {
                        return con.Query<LotesDTO>(@"SELECT AUTONUM_CNTR, ID_CONTEINER, GENERO, FL_ACRESCIMO, LOTE, BL, " +
                            " ITEM, MARCA, QUANTIDADE, NCM, MERCADORIA, LOCAL, SUB_CLASSE, IMPORTADOR, SHIPPER, DESCRICAO_SIGVIG, DESCRICAO, " +
                            " VOLUME, PESO_BRUTO, PESO_APURADO, EMBALAGEM, IMO, UNDG, PATIO, EMPRESA, FLAG_HUB " +
                            "FROM SGIPA..VW_ROMANEIO_CONSULTA_LOTE_ITENS WHERE LOTE = @lote AND ITEM = @item AND AUTONUM_CNTR = @cntr ORDER BY LOTE DESC", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    }

                }
            }
            catch
            {
                return null;
            }

        }

        public LotesDTO ObterDadosRomaneio(int cntr = 0, int lote = 0, int item = 0, int AUTONUM_ROMANEIO = 0)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("cntr", cntr, direction: ParameterDirection.Input);
                parametros.Add("lote", lote, direction: ParameterDirection.Input);
                parametros.Add("item", item, direction: ParameterDirection.Input);


                return con.Query<LotesDTO>(@"
                      SELECT DISTINCT AUTONUM_ROMANEIO, LOTE, AUTONUM_CNTR, INICIO, DEMURRAGE, TERMINO, SOLICITANTE, CONFERENTE, FISCAL, DESOVA, INSPECAO_MAPA, VOLUME, PESO_APURADO, PESO_BRUTO ,IMO, UNDG, GENERO, ITEM, QUANTIDADE, EMBALAGEM, MERCADORIA, MARCA, OBS, FL_OUTRA_AVARIA, 
DESCR_OUTRA_AVARIA, FL_FALTA, DESCR_FALTA, FL_ACRESCIMO, DESCR_ACRESCIMO, FL_AGUARDANDO_RECONHECIMENTO, DESCR_AGUARDANDO_RECONHECIMENTO, LOCALIZACAO, ID_CONTEINER, LOCAL, SUB_CLASSE, PATIO, EMPRESA, TIPO_CARGA, DESCRICAO, NCM , FL_SEM_AVARIA, FLAG_HUB, AUTONUM_CNTR_DESTINO 
  FROM SGIPA..VW_ROMANEIO_CONSULTA_CADASTRADOS WHERE ((LOTE = @lote AND ITEM = @item AND AUTONUM_CNTR = @cntr) OR (AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO)) ORDER BY LOTE DESC ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }
        

        public string SalvarDados(LoteViewModel dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", dados.AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_CNTR", dados.AUTONUM_CNTR, direction: ParameterDirection.Input);
                parametros.Add("LOTE", dados.LOTE, direction: ParameterDirection.Input);
                parametros.Add("INICIO", dados.INICIO, direction: ParameterDirection.Input);
                parametros.Add("DEMURRAGE", dados.DEMURRAGE, direction: ParameterDirection.Input);
                parametros.Add("CONFERENTE", dados.CONFERENTE, direction: ParameterDirection.Input);
                parametros.Add("SOLICITANTE", dados.SOLICITANTE, direction: ParameterDirection.Input);
                parametros.Add("FISCAL", dados.FISCAL, direction: ParameterDirection.Input);
                parametros.Add("DESOVA", dados.DESOVA, direction: ParameterDirection.Input);
                parametros.Add("INSPECAO_MAPA", dados.INSPECAO_MAPA, direction: ParameterDirection.Input);
                parametros.Add("VOLUME_M3", dados.VOLUME, direction: ParameterDirection.Input);
                parametros.Add("PESO_APURADO", dados.PESO_APURADO, direction: ParameterDirection.Input);
                parametros.Add("PESO_BRUTO", dados.PESO_BRUTO, direction: ParameterDirection.Input);
                parametros.Add("IMO", dados.IMO, direction: ParameterDirection.Input);
                parametros.Add("UNDG", dados.UNDG, direction: ParameterDirection.Input);
                parametros.Add("GENERO", dados.GENERO, direction: ParameterDirection.Input);
                parametros.Add("ITEM", dados.ITEM, direction: ParameterDirection.Input);
                parametros.Add("MARCA", dados.MARCA, direction: ParameterDirection.Input);
                parametros.Add("QUANTIDADE", dados.QUANTIDADE, direction: ParameterDirection.Input);
                parametros.Add("MERCADORIA", dados.MERCADORIA, direction: ParameterDirection.Input);
                parametros.Add("EMBALAGEM", dados.EMBALAGEM, direction: ParameterDirection.Input);
                parametros.Add("LOCALIZACAO", dados.LOCALIZACAO, direction: ParameterDirection.Input);
                parametros.Add("OBS", dados.OBS, direction: ParameterDirection.Input);
                parametros.Add("FL_OUTRA_AVARIA", dados.FL_OUTRA_AVARIA, direction: ParameterDirection.Input);
                parametros.Add("DESCR_OUTRA_AVARIA", dados.DESCR_OUTRA_AVARIA, direction: ParameterDirection.Input);
                parametros.Add("FL_FALTA", dados.FL_FALTA, direction: ParameterDirection.Input);
                parametros.Add("DESCR_FALTA", dados.DESCR_FALTA, direction: ParameterDirection.Input);
                parametros.Add("FL_ACRESCIMO", dados.FL_ACRESCIMO, direction: ParameterDirection.Input);
                parametros.Add("DESCR_ACRESCIMO", dados.DESCR_ACRESCIMO, direction: ParameterDirection.Input);
                parametros.Add("FL_AGUARDANDO_RECONHECIMENTO", dados.FL_AGUARDANDO_RECONHECIMENTO, direction: ParameterDirection.Input);
                parametros.Add("DESCR_AGUARDANDO_RECONHECIMENTO", dados.DESCR_AGUARDANDO_RECONHECIMENTO, direction: ParameterDirection.Input);
                parametros.Add("PATIO", dados.PATIO, direction: ParameterDirection.Input);
                parametros.Add("FL_SEM_AVARIA", dados.FL_SEM_AVARIA, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_CNTR_DESTINO", dados.AUTONUM_CNTR_DESTINO, direction: ParameterDirection.Input);

                try
                {
                    if (dados.AUTONUM_ROMANEIO == 0)
                    {
                        con.Execute(@"INSERT INTO SGIPA..TB_ROMANEIO (LOTE, AUTONUM_CNTR, INICIO, DEMURRAGE, SOLICITANTE, CONFERENTE, FISCAL, DESOVA, INSPECAO_MAPA, VOLUME_M3, PESO_APURADO, PESO_BRUTO ,IMO, UNDG, GENERO, ITEM, QUANTIDADE, EMBALAGEM, MERCADORIA, 
                                        MARCA, OBS, LOCALIZACAO, FL_OUTRA_AVARIA , DESCR_OUTRA_AVARIA , FL_FALTA , DESCR_FALTA , FL_ACRESCIMO , DESCR_ACRESCIMO , FL_AGUARDANDO_RECONHECIMENTO , DT_CADASTRO , PATIO, FL_SEM_AVARIA, AUTONUM_CNTR_DESTINO ) VALUES (@LOTE, @AUTONUM_CNTR, CONVERT(DATETIME, @INICIO, 103), CONVERT(DATETIME, @DEMURRAGE, 103), @SOLICITANTE,
                                        @CONFERENTE, @FISCAL, @DESOVA, CONVERT(DATETIME, @INSPECAO_MAPA, 103), @VOLUME_M3, @PESO_APURADO, @PESO_BRUTO, @IMO, @UNDG, @GENERO, @ITEM, @QUANTIDADE, @EMBALAGEM, @MERCADORIA, @MARCA, @OBS, @LOCALIZACAO, @FL_OUTRA_AVARIA , 
                                        @DESCR_OUTRA_AVARIA , @FL_FALTA , @DESCR_FALTA , @FL_ACRESCIMO , @DESCR_ACRESCIMO , @FL_AGUARDANDO_RECONHECIMENTO, GETDATE(), @PATIO, @FL_SEM_AVARIA, @AUTONUM_CNTR_DESTINO ) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());
                    }
                    else
                    {
                        con.Execute(@"UPDATE SGIPA..TB_ROMANEIO SET PATIO = @PATIO, DEMURRAGE = CONVERT(DATETIME, @DEMURRAGE, 103), SOLICITANTE = @SOLICITANTE, CONFERENTE = @CONFERENTE, FISCAL = @FISCAL, DESOVA = @DESOVA, INSPECAO_MAPA = CONVERT(DATETIME, @INSPECAO_MAPA, 103), 
                                        VOLUME_M3 = @VOLUME_M3, PESO_APURADO = @PESO_APURADO, PESO_BRUTO = @PESO_BRUTO ,IMO = @IMO, UNDG = @UNDG, GENERO = @GENERO, ITEM = @ITEM, QUANTIDADE = @QUANTIDADE, EMBALAGEM = @EMBALAGEM, MERCADORIA = @MERCADORIA, 
                                        MARCA = @MARCA, OBS = @OBS , LOCALIZACAO = @LOCALIZACAO, FL_OUTRA_AVARIA = @FL_OUTRA_AVARIA , DESCR_OUTRA_AVARIA = @DESCR_OUTRA_AVARIA , FL_FALTA = @FL_FALTA , DESCR_FALTA = @DESCR_FALTA , FL_ACRESCIMO = @FL_ACRESCIMO , DESCR_ACRESCIMO = @DESCR_ACRESCIMO , 
                                        FL_AGUARDANDO_RECONHECIMENTO = @FL_AGUARDANDO_RECONHECIMENTO , DESCR_AGUARDANDO_RECONHECIMENTO = @DESCR_AGUARDANDO_RECONHECIMENTO , FL_SEM_AVARIA = @FL_SEM_AVARIA, AUTONUM_CNTR_DESTINO = @AUTONUM_CNTR_DESTINO WHERE AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());
                    }


                    return "Informações salvas com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string Finalizar(int romaneio)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", romaneio, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" UPDATE SGIPA..TB_ROMANEIO SET TERMINO = GETDATE() WHERE AUTONUM_CNTR IN (SELECT AUTONUM_CNTR FROM SGIPA..TB_ROMANEIO WHERE AUTONUM_ROMANEIO =  @AUTONUM_ROMANEIO )", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Finalizado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }
        public int VerificarCubagem(int romaneio)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", romaneio, direction: ParameterDirection.Input);
 
                try
                {
                    return con.Query<int>(@" SELECT dbo.FN_ROMANEIO_VERIFICAR_CUBAGEM(@AUTONUM_ROMANEIO) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();


                }
                catch (Exception erro)
                {
                    return 1;
                }
            }
        }

        public string InserirLote(int cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_CNTR", cntr, direction: ParameterDirection.Input);


                try
                {
                    var lote = con.Query<string>(@" INSERT INTO SGIPA..SEQ_BL (DATA) VALUES (GETDATE() ) Select SCOPE_IDENTITY() as ID_BL", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                    con.Query<string>(@" INSERT INTO SGIPA..TB_BL (AUTONUM, DT_BL, FLAG_ACRESCIMO, NUMERO) VALUES ( " + lote + " , GETDATE(), 1, 'ACRESCIMO" + lote + "' ) Select SCOPE_IDENTITY() as ID_BL", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    con.Query<string>(@" INSERT INTO SGIPA..TB_AMR_CNTR_BL (CNTR, BL) VALUES ( " + cntr + " ," + lote + " )", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    con.Query<string>(@" INSERT INTO SGIPA..TB_CARGA_CNTR (BL,ID_CONTEINER, ITEM) VALUES ( " + lote + ", (SELECT ID_CONTEINER FROM TB_CNTR_BL WHERE AUTONUM = " + cntr + "), 1 )", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                    return lote;

                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public IEnumerable<TipoFotoDTO> ListaTipoFoto()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT ID AS AUTONUM, DESCRICAO FROM SGIPA.dbo.TB_TIPOS_PROCESSO ORDER BY DESCRICAO ");

                var query = con.Query<TipoFotoDTO>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public string InsereNovaFoto(FotosRomaneio dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", dados.AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("ID_TIPO_PROCESSO", dados.ID_TIPO_PROCESSO, direction: ParameterDirection.Input);
                parametros.Add("ID_TIPO_FOTO", dados.ID_TIPO_FOTO, direction: ParameterDirection.Input);
                parametros.Add("FOTO", dados.FOTO, direction: ParameterDirection.Input);
                parametros.Add("DESCRICAO", dados.DESCRICAO, direction: ParameterDirection.Input);
                parametros.Add("OBSERVACAO", dados.OBSERVACAO, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@"INSERT INTO SGIPA..TB_FOTO_PROCESSO ( ID_TIPO_FOTO, ID_TIPO_PROCESSO, IMAGEM_PATH, DESCRICAO, OBSERVACAO, AUTONUM_ROMANEIO, DATA_CADASTRO) VALUES ( @ID_TIPO_FOTO, @ID_TIPO_PROCESSO, @FOTO, @DESCRICAO, @OBSERVACAO, @AUTONUM_ROMANEIO, GETDATE()) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Foto cadastrada com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public IEnumerable<FotosRomaneio> ConsultarFotosRomaneio(int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();

                sb.AppendLine("SELECT A.ID AS ID_FOTO, A.AUTONUM_ROMANEIO, A.DESCRICAO, A.OBSERVACAO, B.DESCRICAO AS DESCR_TIPO, A.ID_TIPO_PROCESSO, ID_TIPO_FOTO ");
                sb.AppendLine("FROM SGIPA..TB_FOTO_PROCESSO A ");
                sb.AppendLine("INNER JOIN SGIPA..TB_TIPOS_FOTO B ON A.ID_TIPO_FOTO = B.ID ");
                sb.AppendLine("WHERE A.AUTONUM_ROMANEIO = " + AUTONUM_ROMANEIO + " ");
                sb.AppendLine("ORDER BY DATA_CADASTRO DESC ");
                var query = con.Query<FotosRomaneio>(sb.ToString()).AsEnumerable();

                return query;

            }
        }

        public string BuscarFoto(int ID_FOTO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("ID_FOTO", ID_FOTO, direction: ParameterDirection.Input);

                try
                {
                    return con.Query<string>(@"SELECT IMAGEM_PATH FROM SGIPA..TB_FOTO_PROCESSO WHERE ID = @ID_FOTO", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }
        public string ExcluirFoto(int ID_FOTO, int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("ID_FOTO", ID_FOTO, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" DELETE FROM sgipa..TB_FOTO_PROCESSO WHERE ID = @ID_FOTO AND AUTONUM_ROMANEIO=@AUTONUM_ROMANEIO", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Foto deletada com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }



        public IEnumerable<AvariasRomaneio> ConsultarAvariasRomaneio(int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.AppendLine(" SELECT  ");
                sb.AppendLine(" A.AUTONUM_ROMANEIO, ");
                sb.AppendLine(" A.AUTONUM_ROMANEIO_AVARIAS, ");
                sb.AppendLine(" A.TIPO, ");
                sb.AppendLine(" T.DESCR AS DESCR_TIPO, ");
                sb.AppendLine(" A.LOCAL, ");
                sb.AppendLine(" L.DESCR AS DESCR_LOCAL, ");
                sb.AppendLine(" A.COMPLEMENTO, ");
                sb.AppendLine(" C.DESCR AS DESCR_COMPLEMENTO, ");
                sb.AppendLine(" A.PROVIDENCIA, ");
                sb.AppendLine(" P.DESCR AS DESCR_PROVIDENCIA, ");
                sb.AppendLine(" A.QTD, ");
                sb.AppendLine(" A.PESO, ");
                sb.AppendLine(" A.COMPLEMENTO_IPA, ");
                sb.AppendLine(" A.DT_AVARIA ");
                sb.AppendLine(" FROM SGIPA..TB_ROMANEIO_AVARIAS A ");
                sb.AppendLine(" INNER JOIN SGIPA..DTE_TB_AVARIAS T on T.CODE = A.TIPO AND T.IDENT = 'T' ");
                sb.AppendLine(" INNER JOIN SGIPA..DTE_TB_AVARIAS C on C.CODE = A.COMPLEMENTO AND C.IDENT = 'C' ");
                sb.AppendLine(" INNER JOIN SGIPA..DTE_TB_AVARIAS L on L.CODE = A.LOCAL AND L.IDENT = 'L' ");
                sb.AppendLine(" INNER JOIN SGIPA..DTE_TB_PROVIDENCIAS P on  P.CODE = A.PROVIDENCIA ");
                sb.AppendLine(" WHERE A.AUTONUM_ROMANEIO = " + AUTONUM_ROMANEIO + " ");
                sb.AppendLine(" ORDER BY A.DT_AVARIA DESC  ");

                var query = con.Query<AvariasRomaneio>(sb.ToString()).AsEnumerable();

                return query;

            }
        }





        public string ExcluirAvaria(int AUTONUM_ROMANEIO_AVARIAS, int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO_AVARIAS", AUTONUM_ROMANEIO_AVARIAS, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" DELETE FROM SGIPA..TB_ROMANEIO_AVARIAS WHERE AUTONUM_ROMANEIO_AVARIAS = @AUTONUM_ROMANEIO_AVARIAS AND AUTONUM_ROMANEIO=@AUTONUM_ROMANEIO", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Avaria deletada com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string ExcluirTodaAvaria(int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" DELETE FROM SGIPA..TB_ROMANEIO_AVARIAS WHERE AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Avarias deletadas com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string InsereNovaAvaria(AvariasRomaneio dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", dados.AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("TIPO", dados.TIPO, direction: ParameterDirection.Input);
                parametros.Add("LOCAL", dados.LOCAL, direction: ParameterDirection.Input);
                parametros.Add("COMPLEMENTO", dados.COMPLEMENTO, direction: ParameterDirection.Input);
                parametros.Add("PROVIDENCIA", dados.PROVIDENCIA, direction: ParameterDirection.Input);
                parametros.Add("QTD", dados.QTD, direction: ParameterDirection.Input);
                parametros.Add("PESO", dados.PESO, direction: ParameterDirection.Input);
                parametros.Add("COMPLEMENTO_IPA", dados.COMPLEMENTO_IPA, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@"INSERT INTO SGIPA..TB_ROMANEIO_AVARIAS (AUTONUM_ROMANEIO, TIPO, LOCAL, PROVIDENCIA, COMPLEMENTO, QTD, PESO, COMPLEMENTO_IPA, DT_AVARIA) VALUES ( @AUTONUM_ROMANEIO, @TIPO, @LOCAL, @PROVIDENCIA, @COMPLEMENTO, @QTD, @PESO, @COMPLEMENTO_IPA, GETDATE()) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Avaria cadastrada com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }


        public int ContaAvarias(int romaneio)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", romaneio, direction: ParameterDirection.Input);
                try
                {
                    return con.Query<int>(@"
 SELECT COUNT(*)QTD FROM SGIPA..TB_ROMANEIO A
 LEFT JOIN  SGIPA..TB_ROMANEIO_AVARIAS B ON A.AUTONUM_ROMANEIO = B.AUTONUM_ROMANEIO
 WHERE A.AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO AND ((B.AUTONUM_ROMANEIO IS NOT NULL) OR (A.FL_SEM_AVARIA = 1)) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                }
                catch (Exception erro)
                {
                    return 0;
                }
            }
        }






        public IEnumerable<Lacres> CarregaLacresRO(int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT A.AUTONUM, A.AUTONUM_ARMAZENS_IPA, A.LACRE, A.DT_LANCAMENTO, A.DT_INATIVACAO, A.FLAG_ATIVO, B.DESCR AS ID_CONTEINER ");
                sb.Append(" FROM SGIPA..TB_ROMANEIO_LACRES A ");
                sb.Append(" INNER JOIN SGIPA..TB_ARMAZENS_IPA B ON A.AUTONUM_ARMAZENS_IPA = B.AUTONUM ");
                sb.Append(" WHERE A.AUTONUM_ROMANEIO = " + AUTONUM_ROMANEIO + "");

                var query = con.Query<Lacres>(sb.ToString()).AsEnumerable();

                return query;

            }
        }

        public string InsereLacreRO(int AUTONUM_ROMANEIO, int AUTONUM_ARMAZENS_IPA, string LACRE)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_ARMAZENS_IPA", AUTONUM_ARMAZENS_IPA, direction: ParameterDirection.Input);
                parametros.Add("LACRE", LACRE, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" INSERT INTO  SGIPA..TB_ROMANEIO_LACRES (AUTONUM_ARMAZENS_IPA, LACRE, DT_LANCAMENTO, FLAG_ATIVO, AUTONUM_ROMANEIO ) VALUES ( @AUTONUM_ARMAZENS_IPA, @LACRE, GETDATE(), 1, @AUTONUM_ROMANEIO ) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Cadastrado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string ExcluirLacreRO(int AUTONUM_LACRE)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_LACRE", AUTONUM_LACRE, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" DELETE FROM SGIPA..TB_ROMANEIO_LACRES WHERE AUTONUM = @AUTONUM_LACRE ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Deletado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }



    }

}