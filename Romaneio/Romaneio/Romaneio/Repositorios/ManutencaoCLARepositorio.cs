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
    public class ManutencaoCLARepositorio : IManutencaoCLARepositorio
    {
        public IEnumerable<CLA> ListarCLA(int lote, int item)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT A.AUTONUM AS AUTONUM_BREAK_BULK, A.AUTONUM_CS, A.COMPRIMENTO, A.LARGURA, A.ALTURA, A.QUANTIDADE, A.PESO FROM TB_BREAK_BULK A ");
                sb.Append(" LEFT JOIN TB_CARGA_SOLTA B ON A.AUTONUM_CS = B.AUTONUM ");
                sb.Append(" WHERE B.BL = " + lote + " AND B.ITEM = " + item + "  ");

                return con.Query<CLA>(sb.ToString()).AsEnumerable();

            }

        }
        public int ObterAutonumCS(int lote, int item)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("lote", lote, direction: ParameterDirection.Input);
                parametros.Add("item", item, direction: ParameterDirection.Input);

                return con.Query<int>(@" SELECT AUTONUM FROM SGIPA..TB_CARGA_SOLTA WHERE BL = @lote AND ITEM = @item ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }

        public LoteViewModel ObterCLA(int AUTONUM_BREAK_BULK)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT AUTONUM AS AUTONUM_BREAK_BULK, AUTONUM_CS, COMPRIMENTO, LARGURA, ALTURA, QUANTIDADE AS QUANTIDADE_CLA, PESO FROM TB_BREAK_BULK ");
                sb.Append(" WHERE AUTONUM = " + AUTONUM_BREAK_BULK + "  ");

                return con.Query<LoteViewModel>(sb.ToString()).FirstOrDefault();

            }
        }

        public string SalvarDadosCLA(LoteViewModel dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();

                parametros.Add("AUTONUM_CS", dados.AUTONUM_CS, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_BREAK_BULK", dados.AUTONUM_BREAK_BULK, direction: ParameterDirection.Input);
                parametros.Add("COMPRIMENTO", dados.COMPRIMENTO, direction: ParameterDirection.Input);
                parametros.Add("LARGURA", dados.LARGURA, direction: ParameterDirection.Input);
                parametros.Add("ALTURA", dados.ALTURA, direction: ParameterDirection.Input);
                parametros.Add("QUANTIDADE", dados.QUANTIDADE_CLA, direction: ParameterDirection.Input);
                parametros.Add("PESO", dados.PESO, direction: ParameterDirection.Input);


                StringBuilder sb = new StringBuilder();

                sb.Clear();
                if (dados.AUTONUM_BREAK_BULK == 0)
                {
                    sb.Append("INSERT INTO SGIPA..TB_BREAK_BULK (AUTONUM_CS, COMPRIMENTO, LARGURA, ALTURA, QUANTIDADE, PESO) ");
                    sb.Append("VALUES ( @AUTONUM_CS, @COMPRIMENTO, @LARGURA, @ALTURA, @QUANTIDADE, @PESO )");

                }
                else
                {
                    sb.Append("UPDATE SGIPA..TB_BREAK_BULK SET COMPRIMENTO = @COMPRIMENTO, LARGURA = @LARGURA, ALTURA = @ALTURA, QUANTIDADE = @QUANTIDADE, PESO = @PESO WHERE AUTONUM = @AUTONUM_BREAK_BULK");
                }


                con.Query<string>(sb.ToString(), parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                return "Informações salvas com sucesso!";


            }
        }

        public int VerificarCubagem(LoteViewModel dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_CS", dados.AUTONUM_CS, direction: ParameterDirection.Input);
                

                try
                {
                    return con.Query<int>(@" SELECT dbo.FN_VERIFICAR_CUBAGEM_CS(@AUTONUM_CS) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                }
                catch (Exception erro)
                {
                    return 1;
                }
            }
        }

        public string ExcluirCLA(int AUTONUM_BREAK_BULK)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM", AUTONUM_BREAK_BULK, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" DELETE FROM sgipa..TB_BREAK_BULK WHERE AUTONUM = @AUTONUM  ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Registro deletado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }







        public IEnumerable<CLA> ListarCLARomaneio(int AUTONUM_ROMANEIO)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT AUTONUM_ROMANEIO_BREAK_BULK AS AUTONUM_BREAK_BULK, COMPRIMENTO, LARGURA, ALTURA, QUANTIDADE, PESO FROM TB_ROMANEIO_BREAK_BULK ");
                 sb.Append(" WHERE AUTONUM_ROMANEIO = " + AUTONUM_ROMANEIO  );

                return con.Query<CLA>(sb.ToString()).AsEnumerable();

            }

        }
 
        public LoteViewModel ObterCLARomaneio(int AUTONUM_BREAK_BULK)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT AUTONUM_ROMANEIO_BREAK_BULK AS AUTONUM_BREAK_BULK, COMPRIMENTO, LARGURA, ALTURA, QUANTIDADE AS QUANTIDADE_CLA, PESO FROM TB_ROMANEIO_BREAK_BULK ");
                sb.Append(" WHERE AUTONUM_ROMANEIO_BREAK_BULK = " + AUTONUM_BREAK_BULK + "  ");

                return con.Query<LoteViewModel>(sb.ToString()).FirstOrDefault();

            }
        }

        public string SalvarDadosCLARomaneio(LoteViewModel dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();

                parametros.Add("AUTONUM_ROMANEIO", dados.AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_BREAK_BULK", dados.AUTONUM_BREAK_BULK, direction: ParameterDirection.Input);
                parametros.Add("COMPRIMENTO", dados.COMPRIMENTO, direction: ParameterDirection.Input);
                parametros.Add("LARGURA", dados.LARGURA, direction: ParameterDirection.Input);
                parametros.Add("ALTURA", dados.ALTURA, direction: ParameterDirection.Input);
                parametros.Add("QUANTIDADE", dados.QUANTIDADE_CLA, direction: ParameterDirection.Input);
                parametros.Add("PESO", dados.PESO, direction: ParameterDirection.Input);


                StringBuilder sb = new StringBuilder();

                sb.Clear();
                if (dados.AUTONUM_BREAK_BULK == 0)
                {
                    sb.Append("INSERT INTO SGIPA..TB_ROMANEIO_BREAK_BULK (AUTONUM_ROMANEIO, COMPRIMENTO, LARGURA, ALTURA, QUANTIDADE, PESO) ");
                    sb.Append("VALUES ( @AUTONUM_ROMANEIO, @COMPRIMENTO, @LARGURA, @ALTURA, @QUANTIDADE, @PESO )");

                }
                else
                {
                    sb.Append("UPDATE SGIPA..TB_ROMANEIO_BREAK_BULK SET COMPRIMENTO = @COMPRIMENTO, LARGURA = @LARGURA, ALTURA = @ALTURA, QUANTIDADE = @QUANTIDADE, PESO = @PESO WHERE AUTONUM_ROMANEIO_BREAK_BULK = @AUTONUM_BREAK_BULK");
                }


                con.Query<string>(sb.ToString(), parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                return "Informações salvas com sucesso!";


            }
        }
        public string ExcluirCLARomaneio(int AUTONUM_BREAK_BULK)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM", AUTONUM_BREAK_BULK, direction: ParameterDirection.Input);

                try
                {
                    con.Execute(@" DELETE FROM TB_ROMANEIO_BREAK_BULK WHERE AUTONUM_ROMANEIO_BREAK_BULK = @AUTONUM  ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Registro deletado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }
    }

}