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
    public class FechamentosRepositorio : IFechamentosRepositorio
    {
        public IEnumerable<Lotes> ListaLotesFechados(int cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT LOTE AS AUTONUM, LOTE FROM VW_ROMANEIO_FECHAMENTO_LOTE_CNTR WHERE AUTONUM_CNTR = " + cntr + " ORDER BY LOTE ");

                var query = con.Query<Lotes>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<Conteiners> ConsultarFechamentos(int PATIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();

                sb.AppendLine("SELECT DISTINCT AUTONUM_CNTR, ID_CONTEINER, DESCR_TIPO, TAMANHO, DT_ENTRADA, DT_AGENDAMENTO, ");
                sb.AppendLine("LACRE_ORIGEM, LACRE_IPA, VIAGEM, BOL_DESUNIT, DDA, SOLICITADO, DESCRICAO_SIGVIG, POSICAO, ");
                sb.AppendLine("MIN(ABERTURA) AS ABERTURA, MAX(FECHAMENTO) AS FECHAMENTO ");
                sb.AppendLine("FROM VW_ROMANEIO_FECHAMENTO_LOTE_CNTR WHERE PATIO = " + PATIO + " ");
                sb.AppendLine("GROUP BY AUTONUM_CNTR, ID_CONTEINER, DESCR_TIPO, TAMANHO, DT_ENTRADA, DT_AGENDAMENTO,  LACRE_ORIGEM, LACRE_IPA, VIAGEM, ");
                sb.AppendLine("BOL_DESUNIT, DDA, SOLICITADO, DESCRICAO_SIGVIG, POSICAO ");
                sb.AppendLine("ORDER BY DT_ENTRADA DESC ");

                var query = con.Query<Conteiners>(sb.ToString()).AsEnumerable();

                return query;

            }
        }
        public LotesDTO ObterDadosRomaneioFechado(int cntr = 0, int lote = 0, int item = 0, int AUTONUM_ROMANEIO = 0)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("cntr", cntr, direction: ParameterDirection.Input);
                parametros.Add("lote", lote, direction: ParameterDirection.Input);
                parametros.Add("item", item, direction: ParameterDirection.Input);

                if (item == 0)
                {
                    return con.Query<LotesDTO>(@"
                      SELECT DISTINCT AUTONUM_ROMANEIO, LOTE, AUTONUM_CNTR, INICIO, DEMURRAGE, TERMINO, SOLICITANTE, CONFERENTE, FISCAL, DESOVA, INSPECAO_MAPA, VOLUME, PESO_APURADO, PESO_BRUTO ,IMO, UNDG, GENERO, ITEM, QUANTIDADE, EMBALAGEM, MERCADORIA, MARCA, OBS, FL_OUTRA_AVARIA, 
DESCR_OUTRA_AVARIA, FL_FALTA, DESCR_FALTA, FL_ACRESCIMO, DESCR_ACRESCIMO, FL_AGUARDANDO_RECONHECIMENTO, DESCR_AGUARDANDO_RECONHECIMENTO, LOCALIZACAO, ID_CONTEINER, LOCAL, SUB_CLASSE, PATIO, EMPRESA, TIPO_CARGA, DESCRICAO, NCM, FL_SEM_AVARIA 
  FROM SGIPA..VW_ROMANEIO_CONSULTA_CADASTRADOS WHERE ((LOTE = @lote AND AUTONUM_CNTR = @cntr) OR (AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO)) ORDER BY LOTE DESC ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                }
                else
                {
                    return con.Query<LotesDTO>(@"
                      SELECT DISTINCT AUTONUM_ROMANEIO, LOTE, AUTONUM_CNTR, INICIO, DEMURRAGE, TERMINO, SOLICITANTE, CONFERENTE, FISCAL, DESOVA, INSPECAO_MAPA, VOLUME, PESO_APURADO, PESO_BRUTO ,IMO, UNDG, GENERO, ITEM, QUANTIDADE, EMBALAGEM, MERCADORIA, MARCA, OBS, FL_OUTRA_AVARIA, 
DESCR_OUTRA_AVARIA, FL_FALTA, DESCR_FALTA, FL_ACRESCIMO, DESCR_ACRESCIMO, FL_AGUARDANDO_RECONHECIMENTO, DESCR_AGUARDANDO_RECONHECIMENTO, LOCALIZACAO, ID_CONTEINER, LOCAL, SUB_CLASSE, PATIO, EMPRESA, TIPO_CARGA, DESCRICAO, NCM , FL_SEM_AVARIA 
  FROM SGIPA..VW_ROMANEIO_CONSULTA_CADASTRADOS WHERE ((LOTE = @lote AND ITEM = @item AND AUTONUM_CNTR = @cntr) OR (AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO)) ORDER BY LOTE DESC ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                }
            }
        }
        public IEnumerable<LotesDTO> ConsultarFechamentosLotesConteiners(string cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.AppendLine("SELECT DISTINCT AUTONUM_ROMANEIO, AUTONUM_CNTR, LOTE, BL, ITEM, MARCA, QUANTIDADE, NCM, MERCADORIA, LOCALIZACAO, IMPORTADOR, SHIPPER   ");
                sb.AppendLine("FROM VW_ROMANEIO_FECHAMENTO_LOTE_ITENS ");
                sb.AppendLine("WHERE AUTONUM_CNTR = " + cntr);
                sb.AppendLine(" ORDER BY LOTE DESC ");
                var query = con.Query<LotesDTO>(sb.ToString()).AsEnumerable();

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
                        return con.Query<LotesDTO>(@"SELECT AUTONUM_CNTR, ID_CONTEINER, LOTE FROM VW_ROMANEIO_CONSULTA_LOTE_CNTR WHERE LOTE = @lote AND AUTONUM_CNTR = @cntr ORDER BY LOTE DESC", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    }
                    else
                    {
                        return con.Query<LotesDTO>(@"SELECT AUTONUM_CNTR, ID_CONTEINER, GENERO, FL_ACRESCIMO, LOTE, BL, " +
                            " ITEM, MARCA, QUANTIDADE, NCM, MERCADORIA, LOCAL, SUB_CLASSE, IMPORTADOR, SHIPPER, DESCRICAO_SIGVIG, DESCRICAO, " +
                            " VOLUME, PESO_BRUTO, PESO_APURADO, EMBALAGEM, IMO, UNDG, PATIO, EMPRESA " +
                            "FROM SGIPA..VW_ROMANEIO_CONSULTA_LOTE_ITENS WHERE LOTE = @lote AND ITEM = @item AND AUTONUM_CNTR = @cntr ORDER BY LOTE DESC", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
                    }


                }
            }
            catch
            {
                return null;
            }


        }



    }

}