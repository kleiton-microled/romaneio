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
using Romaneio.Models.ViewModels;

namespace Romaneio.Repositorios
{
    public class CompartilhaScannerRepositorio : ICompartilhaScannerRepositorio
    {


        public IEnumerable<Destinos> ConsultarDestinos()

        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append("SELECT AUTONUM, RAZAO AS DESCRICAO FROM SGIPA..TB_CAD_PARCEIROS WHERE FLAG_PARCEIRO_SCANNER = 1 AND FLAG_ATIVO = 1 ORDER BY RAZAO ");

                var query = con.Query<Destinos>(sb.ToString()).AsEnumerable();

                return query;


            }
        }

        public IEnumerable<AutorizacaoSaida> ConsultarDados(int patio, string viagem, string cntr, string status)

        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT A.SISTEMA, A.AUTONUM, A.ID_CONTEINER, A.TAMANHO, A.EF, A.VIAGEM_CODESP AS VIAGEM, ");
                sb.Append(" A.TERMINAL, C.RAZAO AS DESTINO, A.DT_SAIDA_SCANNER ");
                sb.Append(" FROM  OPERADOR..VW_AG_CNTR_SCANNER_COMPARTILHADO A ");
                sb.Append(" LEFT JOIN SGIPA..TB_CAD_PARCEIROS C ON A.PARCEIRO_DESTINO_SCANNER = C.AUTONUM ");

                if (status == "Autorizacao")
                {
                    sb.Append(" WHERE A.DT_SCANNEAMENTO IS NULL AND A.PARCEIRO_DESTINO_SCANNER IS NULL ");

                }
                else if(status == "Saida")
                {
                    sb.Append(" WHERE A.PARCEIRO_DESTINO_SCANNER IS NOT NULL AND A.DT_SAIDA_SCANNER IS NULL ");

                }
                else 
                {
                    sb.Append(" WHERE A.DT_SAIDA_SCANNER IS NOT NULL AND A.DT_RETORNO_SCANNER IS NULL ");

                }


 
                if (cntr != "")
                {
                    sb.Append(" AND A.ID_CONTEINER ='" + cntr +"' ");

                }

                if (viagem != "")
                {
                    sb.Append("AND A.VIAGEM_CODESP ='" + viagem + "' ");

                }
 
                if( patio > 0)
                {
                    sb.Append(" AND A.PATIO =" + patio );

                }
 


                var query = con.Query<AutorizacaoSaida>(sb.ToString()).AsEnumerable();

                return query;


            }
        }

        public string Autorizar(CompartilhaScannerViewModel dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM", dados.AUTONUM, direction: ParameterDirection.Input);
                parametros.Add("ID_CONTEINER", dados.ID_CONTEINER, direction: ParameterDirection.Input);
                parametros.Add("PARCEIRO_DESTINO", dados.DESTINO, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_GATE", dados.AUTONUM_GATE, direction: ParameterDirection.Input);
 
                try
                {
                    StringBuilder sb = new StringBuilder();


                    sb.Clear();


                    sb.AppendLine("INSERT INTO SGIPA..TB_COMPARTILHA_SCANNER (DT_CADASTRO,PARCEIRO_DESTINO,AUTONUM_GATE,ID_CONTEINER,AUTONUM_CNTR_BL,AUTONUM_PATIO,AUTONUM_CT) ");
                    sb.AppendLine("VALUES(GETDATE(), @PARCEIRO_DESTINO, @AUTONUM_GATE, @ID_CONTEINER, ");

                    if (dados.SISTEMA == "I")
                    {
                        sb.AppendLine(" @AUTONUM, 0, 0 ");

                    }
                    else if (dados.SISTEMA == "R")
                    {
                        sb.AppendLine(" 0, @AUTONUM, 0  ");

                    }
                    else //A
                    {
                        sb.AppendLine(" 0 , 0 , @AUTONUM ");

                    }
                    sb.AppendLine(" ) ");

                    con.Execute(sb.ToString(), parametros, commandTimeout: Config.QueryTimeoutInSeconds());


                    return "Informações salvas com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }


    }

}