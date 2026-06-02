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
    public class MovimentacaoCSRepositorio : IMovimentacaoCSRepositorio
    {
        public Marcante ConsultarMarcante(string marcante)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT AUTONUM,AUTONUM_CARGA,AUTONUM_CARGA_AC,VOLUMES,AUTONUM_CS_YARD,AUTONUM_CS_YARD_AC, ");
                sb.Append(" ID_CONTEINER_IMPRESSO, ARMAZEM_REC, YARD_REC, AUTONUM_CEXP, PATIO,FLAG_RECONHECIMENTO,LOTE_IMPRESSO ");
                sb.Append(" FROM sgipa.dbo.fnInvent_Sistemas_Arm_Col_patio('" + marcante + "')");
                 return con.Query<Marcante>(sb.ToString(), commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }
        public Marcante ConsultarMarcanteInventArmazem(string marcante)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("SELECT LOTE,BL,EMBALAGEM,IMPORTADOR,NAVIO,QTDE as QUANTIDADE_MARCANTE,AUTONUMCS,MERCADORIA,DATA_ENTRADA,TIPO_DOC,DESCR_ARMAZEM,POSICAO,ID_GRAVACAO ");
                sb.Append(" FROM OPERADOR..VW_INVENT_ARMAZEM WHERE MARCANTE='" + marcante + "' ");
                return con.Query<Marcante>(sb.ToString(), commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }

        public string ConsultarArmazem(int armazem_rec)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT DESCR, FLAG_CT FROM SGIPA..TB_ARMAZENS_IPA WHERE AUTONUM = " + armazem_rec + " ");

                return con.Query<string>(sb.ToString(), commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }
        public MovimentacaoCSViewModel ConsultarCS(int autonum_cs)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT BL, ITEM FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM = " + autonum_cs + " ");

                return con.Query<MovimentacaoCSViewModel>(sb.ToString(), commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }
        
        public MovimentacaoCSViewModel ConsultarLote(int lote, string marcante)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT MERCADORIA, MARCA, ENTRADA AS ENT_DESOVA, CNTR_DESOVA AS ID_CONTEINER, IMO, MOTIVO_PROX_MVTO AS MOV_AGEND, IMPORTADOR FROM sgipa.dbo.fnInvent_Sistemas_Arm_CollOTE(" + lote + ",'" + marcante + "')");

                return con.Query<MovimentacaoCSViewModel>(sb.ToString(), commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }

        public IEnumerable<Itens> ConsultarItens(int lote, string marcante, int autonum_cs_yard)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT AUTONUM AS ID_GRAVACAO, DISPLAY AS DESCR FROM sgipa.dbo.fnInvent_Sistemas_Arm_ColItem(" + lote + ",'" + marcante + "'," + autonum_cs_yard + ")");
                var query = con.Query<Itens>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public MovimentacaoCSViewModel CarregaDadosItem(string id_Gravacao)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();

                sb.Append("SELECT  QTDE as QTD_EMBALAGEM, EMBALAGEM, (DESCR_ARMAZEM + ' ' + POSICAO) AS LOCAL ");
                sb.Append(" FROM OPERADOR..VW_INVENT_ARMAZEM");
                sb.Append(" WHERE ID_GRAVACAO='" + id_Gravacao + "'");

                return con.Query<MovimentacaoCSViewModel>(sb.ToString(), commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();


            }

        }
        public IEnumerable<Motivo> ConsultarMotivo() 
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT AUTONUM, DESCRICAO AS DESCR  FROM OPERADOR..TB_CAD_MOTIVO WHERE FLAG_ATIVO = 1 ORDER BY DESCRICAO ");
                var query = con.Query<Motivo>(sb.ToString()).AsEnumerable();

                return query;


            }

        }


        public string SalvarDados(MovimentacaoCSViewModel dados)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_CS", dados.AUTONUM_CS_YARD, direction: ParameterDirection.Input);
                parametros.Add("ARMAZEM", dados.ARMAZEM, direction: ParameterDirection.Input);
                parametros.Add("YARD", dados.LOCAL, direction: ParameterDirection.Input);
                parametros.Add("QUANTIDADE", dados.QTD_LOCAL, direction: ParameterDirection.Input);
                parametros.Add("MOTIVO", dados.MOTIVO, direction: ParameterDirection.Input);
                parametros.Add("FL_FRENTE", dados.FL_FRENTE, direction: ParameterDirection.Input);
                parametros.Add("FL_FUNDO", dados.FL_FUNDO, direction: ParameterDirection.Input);
                parametros.Add("FL_LE", dados.FL_LE, direction: ParameterDirection.Input);
                parametros.Add("FL_LD", dados.FL_LD, direction: ParameterDirection.Input);
                parametros.Add("MARCANTE", dados.MARCANTE, direction: ParameterDirection.Input);

                //TB_CARGA_SOLTA_YARD_AC
                //TB_CARGA_SOLTA_YARD_CEXP
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD) ");
                sb.Append("VALUES ( @AUTONUM_CS, @ARMAZEM, @YARD, 'I', @QUANTIDADE, @MOTIVO, 0, @FL_FRENTE, @FL_FUNDO, @FL_LE, @FL_LD )");

                con.Query<string>(sb.ToString(), parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

                sb.Clear();
                sb.Append("INSERT INTO SGIPA..TB_HIST_SHIFTING_CS (MARCANTE, ARMAZEM, YARD, DT_MOV) ");
                sb.Append("VALUES ( @MARCANTE, @ARMAZEM, @YARD, GETDATE() )");

                con.Query<string>(sb.ToString(), parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();



                return "Informações salvas com sucesso!";
            }
        }

        public IEnumerable<Historico> Carrega_Hist(string marcante)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT ISNULL(B.DESCR, '')ARMAZEM, upper(isnull(A.YARD, ''))LOCAL, rtrim(ltrim(convert(char, A.DT_MOV, 3))) + ' ' + left(rtrim(ltrim(convert(char, A.DT_MOV, 108))), 5)DATA, U.USUARIO ");
                sb.Append(" FROM SGIPA..TB_HIST_SHIFTING_CS A LEFT JOIN SGIPA..TB_ARMAZENS_IPA B ");
                sb.Append(" ON A.ARMAZEM=B.AUTONUM ");
                sb.Append(" LEFT JOIN SGIPA..TB_CAD_USUARIOS U ON A.USUARIO=U.AUTONUM ");
                sb.Append(" WHERE A.MARCANTE = " + marcante);
                sb.Append(" ORDER BY A.DT_MOV DESC ");
                var query = con.Query<Historico>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public IEnumerable<CargasCT> CarregaCargasCT(string armazem)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT DESCR_ARMAZEM,FINALITY,BL,LOTE,QTDE,EMBALAGEM,MERCADORIA,MARCA,IMPORTADOR,NAVIO,VIAGEM,DATA_ENTRADA,DATA_DESOVA,CNTR_DESOVA, ");
                sb.Append(" PESO, TIPO_DOC, REPLACE(IMO,'0','') AS IMO, DIAS_ARM_FAIXA, MARCANTE  ");
                sb.Append(" FROM OPERADOR..VW_INVENT_ARMAZEM  ");
                sb.Append(" WHERE DESCR_ARMAZEM = '" + armazem + "'");
                sb.Append(" ORDER BY LOTE ");
                var query = con.Query<CargasCT>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public IEnumerable<Lacres> CarregaLacresCT(int autonum_armazens_ipa )
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append(" SELECT AUTONUM, AUTONUM_ARMAZENS_IPA, LACRE, DT_LANCAMENTO, DT_INATIVACAO, FLAG_ATIVO ");
                sb.Append(" FROM SGIPA..TB_LACRES_ARMAZENS_IPA  ");
                sb.Append(" WHERE AUTONUM_ARMAZENS_IPA = " + autonum_armazens_ipa + "");
                var query = con.Query<Lacres>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public string InsereLacreCT(int AUTONUM_ARMAZENS_IPA, string LACRE)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ARMAZENS_IPA", AUTONUM_ARMAZENS_IPA, direction: ParameterDirection.Input);
                parametros.Add("LACRE", LACRE, direction: ParameterDirection.Input);
 
                try
                {
                    con.Execute(@"INSERT INTO SGIPA..TB_LACRES_ARMAZENS_IPA (AUTONUM_ARMAZENS_IPA, LACRE, DT_LANCAMENTO, FLAG_ATIVO ) VALUES ( @AUTONUM_ARMAZENS_IPA, @LACRE, GETDATE(), 1) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "Cadastrado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string ExcluirLacreCT(int AUTONUM_LACRE)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_LACRE", AUTONUM_LACRE, direction: ParameterDirection.Input);
 
                try
                {
                    con.Execute(@" DELETE FROM SGIPA..TB_LACRES_ARMAZENS_IPA WHERE AUTONUM = @AUTONUM_LACRE ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

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
