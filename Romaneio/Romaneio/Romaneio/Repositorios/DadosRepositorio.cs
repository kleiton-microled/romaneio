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
    public class DadosRepositorio : IDadosRepositorio
    {
        public IEnumerable<TiposPatios> ConsultarPatios()

        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append("SELECT AUTONUM, DESCR_RESUMIDO AS DESCRICAO FROM OPERADOR..TB_PATIOS WHERE FLAG_ATIVO = 1 ORDER BY DESCR_RESUMIDO");

                var query = con.Query<TiposPatios>(sb.ToString()).AsEnumerable();

                return query;


            }
        }

     
        public Login ConsultarUsuario(string usuario)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("usuario", usuario, direction: ParameterDirection.Input);
                return con.Query<Login>(@"SELECT AUTONUM, USUARIO, SENHA, NOME, CPF, ISNULL(flag_carregamento_marcante_bip, 0) AS FLAG_CARREGAMENTO_MARCANTE_BIP FROM SGIPA..TB_CAD_USUARIOS WHERE USUARIO = @usuario", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }

        public string ConsultarNomeUsuario(string id)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("id", id, direction: ParameterDirection.Input);

                return con.Query<string>(@" SELECT NOME FROM SGIPA..TB_CAD_USUARIOS WHERE AUTONUM = @id", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }

        public IEnumerable<Solicitantes> TiposSolicitantes()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append("SELECT 1 AUTONUM, 'IMPORTADOR' DESCRICAO UNION SELECT 2 AUTONUM, 'ARMADOR' DESCRICAO UNION SELECT 3 AUTONUM, 'NVOCC' DESCRICAO UNION SELECT 4 AUTONUM, 'DESPACHANTE/COMISSARIA' DESCRICAO UNION SELECT 5 AUTONUM, 'ALFANDEGA' DESCRICAO   ORDER BY DESCRICAO ASC");

                var query = con.Query<Solicitantes>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<Embalagens> TiposEmbalagens()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT CODE AS AUTONUM, DESCR AS DESCRICAO, AUTONUM_EMB FROM  SGIPA..DTE_TB_EMBALAGENS ORDER BY DESCR ");

                var query = con.Query<Embalagens>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<TipoDesova> TiposDesova()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append("SELECT 1 AUTONUM, 'DESOVA MANUAL' DESCRICAO UNION SELECT 2 AUTONUM, 'DESOVA MECANIZADA' DESCRICAO UNION SELECT 3 AUTONUM, 'VISTORIA MAPA' DESCRICAO ORDER BY DESCRICAO ");

                var query = con.Query<TipoDesova>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<Generos> TiposGeneros()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append("SELECT AUTONUM, DESCR AS DESCRICAO FROM SGIPA..TB_CAD_GRUPO_PRODUTOS ORDER BY DESCR ASC");

                var query = con.Query<Generos>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public IEnumerable<RelacaoUnidadesViewModel> ConsultarUnidades(string lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();

                sb.AppendLine(" SELECT DISTINCT C.AUTONUM AS LOTE , A.ID_CONTEINER ");
                sb.AppendLine(" FROM SGIPA..TB_CNTR_BL  A ");
                sb.AppendLine(" INNER JOIN SGIPA..TB_AMR_CNTR_BL  B ON B.CNTR = A.AUTONUM ");
                sb.AppendLine(" INNER JOIN SGIPA..TB_BL  C ON C.AUTONUM = B.BL ");
                sb.AppendLine("  WHERE C.AUTONUM = " + lote);

                var query = con.Query<RelacaoUnidadesViewModel>(sb.ToString()).AsEnumerable();

                return query;

            }
        }
        public IEnumerable<NCMViewModel> ConsultarNCM(string lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();

                sb.AppendLine(" SELECT DISTINCT UN, ONU, NCM, DESCR_NCM ");
                sb.AppendLine(" FROM SGIPA..VW_NCM_CE A");
                sb.AppendLine(" INNER JOIN SGIPA..TB_BL B ON B.CE = A.CE ");
                sb.AppendLine("  WHERE B.AUTONUM = " + lote);

                var query = con.Query<NCMViewModel>(sb.ToString()).AsEnumerable();

                return query;

            }
        }
        public DescricaoDTO ConsultarDescr(string lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("LOTE", lote, direction: ParameterDirection.Input);
                return con.Query<DescricaoDTO>(@" SELECT ISNULL(MERCADORIADESC1,'') + ISNULL(MERCADORIADESC2,'') AS Descricao FROM SGIPA..TB_DESCONSOLIDACAO_CE A INNER JOIN SGIPA..TB_BL B ON A.CE = B.CE WHERE B.AUTONUM = @LOTE ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }
        public LocalizacaoDTO ConsultarLocalizacao(string lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("LOTE", lote, direction: ParameterDirection.Input);
                return con.Query<LocalizacaoDTO>(@" SELECT B.DESCR as SubClasse, B.LOCAL AS Local FROM TB_BL A INNER JOIN TB_TIPO_SUB_CLASSE B ON B.AUTONUM = A.SUB_CLASSE WHERE A.AUTONUM = @LOTE ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }

        public IEnumerable<Posicionamento> ConsultarPosicoesItem(string patio, string cntr, string lote, string item)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.AppendLine(" SELECT DISTINCT A.AUTONUM AS AUTONUM_YARD, A.AUTONUM_CS, CS.ITEM, B.DESCR AS ARMAZEM, A.YARD AS POSICAO, A.QUANTIDADE ");
                sb.AppendLine(" FROM SGIPA..TB_CARGA_SOLTA_YARD A ");
                sb.AppendLine(" LEFT JOIN  SGIPA..TB_ARMAZENS_IPA B ON A.ARMAZEM = B.AUTONUM ");
                sb.AppendLine(" INNER JOIN  SGIPA..TB_CARGA_SOLTA CS ON A.AUTONUM_CS = CS.AUTONUM ");
                sb.AppendLine(" WHERE ITEM = " + item + " AND CNTR = " + cntr + " AND BL =  " + lote + "  AND CS.PATIO= " + patio); //+ " AND ORIGEM='I' ");
                sb.AppendLine(" ORDER BY A.YARD ");
                var query = con.Query<Posicionamento>(sb.ToString()).AsEnumerable();

                return query;

            }
        }

        public IEnumerable<ConteinersDTO> ConteinersDestino(int patio)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("SELECT AUTONUM, DESCR AS ID_CONTEINER FROM SGIPA..TB_ARMAZENS_IPA WHERE PATIO= " + patio + " AND DT_ENTRADA IS NOT NULL AND DT_SAIDA IS NULL ORDER BY DESCR ASC");

                return con.Query<ConteinersDTO>(sb.ToString()).AsEnumerable();

            }

        }
        public IEnumerable<Armazens> ListaArmazens(string patio)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                 sb.Append(" SELECT DISTINCT DESCR AS DESCRICAO , AUTONUM, FLAG_CT FROM SGIPA..TB_ARMAZENS_IPA  WHERE PATIO= " + patio + " ORDER BY DESCRICAO  ");

                var query = con.Query<Armazens>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<Armazens> ListaArmazensAbertos(string patio)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT DESCR AS DESCRICAO , AUTONUM, FLAG_CT FROM SGIPA..TB_ARMAZENS_IPA  WHERE PATIO= " + patio + " AND AUTONUM NOT IN (SELECT ISNULL(AUTONUM_ARMAZENS_IPA,0) FROM TB_LACRES_ARMAZENS_IPA WHERE FLAG_ATIVO = 1) ORDER BY DESCRICAO  ");

                var query = con.Query<Armazens>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<Itens> ListaItens(int lote, int cntr)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT A.ITEM, 'ITEM: ' + CONVERT(VARCHAR,A.ITEM) + ' -  QTD: ' + CONVERT(VARCHAR,A.QUANTIDADE) + ' - ' + CONVERT(VARCHAR,A.MERCADORIA) AS DESCR ");
                sb.Append(" FROM SGIPA..TB_CARGA_CNTR A ");
                sb.Append(" INNER JOIN TB_CNTR_BL B ON A.ID_CONTEINER =  B.ID_CONTEINER ");
                sb.Append(" WHERE A.BL = " + lote + " ");
                if (cntr > 0)
                {
                    sb.Append(" AND B.AUTONUM = " + cntr + " ");
                }
                sb.Append(" ORDER BY A.ITEM ");

                var query = con.Query<Itens>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public string ConsultarSolicitanteCNTR(string AUTONUM_ROMANEIO, string CNTR )
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("CNTR", CNTR, direction: ParameterDirection.Input);
 
                return con.Query<string>(@" SELECT DISTINCT CASE WHEN ISNULL(A.SOLICITANTE,0) = 0 THEN (SELECT MAX(S.SOLICITANTE) FROM SGIPA..TB_ROMANEIO S WHERE ISNULL(S.SOLICITANTE,0) <> 0 AND A.AUTONUM_CNTR = S.AUTONUM_CNTR) ELSE SOLICITANTE END SOLICITANTE  FROM SGIPA..TB_ROMANEIO A WHERE ((  AUTONUM_CNTR = @CNTR) OR (AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO)) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }

        public string ConsultarDesovaCNTR(string AUTONUM_ROMANEIO, string CNTR)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_ROMANEIO", AUTONUM_ROMANEIO, direction: ParameterDirection.Input);
                parametros.Add("CNTR", CNTR, direction: ParameterDirection.Input); 

                return con.Query<string>(@" SELECT DISTINCT CASE WHEN ISNULL(A.DESOVA,0) = 0 THEN (SELECT MAX(S.DESOVA) FROM SGIPA..TB_ROMANEIO S WHERE ISNULL(S.DESOVA,0) <> 0 AND A.AUTONUM_CNTR = S.AUTONUM_CNTR) ELSE DESOVA END DESOVA  FROM SGIPA..TB_ROMANEIO A WHERE (( AUTONUM_CNTR = @CNTR) OR (AUTONUM_ROMANEIO = @AUTONUM_ROMANEIO)) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }
 
        public string ObterAutonumCS(string patio, string cntr, string lote, string item)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("PATIO", patio, direction: ParameterDirection.Input);
                parametros.Add("CNTR", cntr, direction: ParameterDirection.Input);
                parametros.Add("LOTE", lote, direction: ParameterDirection.Input);
                parametros.Add("ITEM", item, direction: ParameterDirection.Input);

                return con.Query<string>(@"SELECT AUTONUM AS AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA WHERE BL = @LOTE AND CNTR = @CNTR AND PATIO = @PATIO AND ITEM = @ITEM ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();
            }
        }


        public string InsereYard(string AUTONUM_CS, string QTD, string ARM, string YARD)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_CS", AUTONUM_CS, direction: ParameterDirection.Input);
                parametros.Add("QTD", QTD, direction: ParameterDirection.Input);
                parametros.Add("ARM", ARM, direction: ParameterDirection.Input);
                parametros.Add("YARD", YARD, direction: ParameterDirection.Input);

                try
                {   //o que preenche o campo ORIGEM desse insert ?
                    con.Execute(@"INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD ( AUTONUM_CS, ARMAZEM, YARD, QUANTIDADE) VALUES (@AUTONUM_CS, @ARM, @YARD, @QTD) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());
                    con.Execute(@"UPDATE SGIPA..TB_CARGA_SOLTA SET LOCAL_ISO9000 = @YARD, ARMAZEM_IPA = @ARM WHERE AUTONUM = @AUTONUM_CS ) ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "YARD cadastrado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string AlteraYard(int AUTONUM_CS, string QTD, string ARM, string YARD)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM_CS", AUTONUM_CS, direction: ParameterDirection.Input);
                parametros.Add("QTD", QTD, direction: ParameterDirection.Input);
                parametros.Add("ARM", ARM, direction: ParameterDirection.Input);
                parametros.Add("YARD", YARD, direction: ParameterDirection.Input);

                try
                {
                    return "YARD alterado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public string DeletaYard(int AUTONUM, int AUTONUM_CS)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("AUTONUM", AUTONUM, direction: ParameterDirection.Input);
                parametros.Add("AUTONUM_CS", AUTONUM_CS, direction: ParameterDirection.Input);

                try
                {   //o que preenche o campo ORIGEM desse insert ?
                    con.Execute(@" DELETE FROM sgipa..TB_CARGA_SOLTA_YARD WHERE AUTONUM = @AUTONUM  ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());
                    con.Execute(@" UPDATE SGIPA TB_CARGA_SOLTA SET LOCAL_ISO9000='', ARMAZEM_IPA = NULL WHERE AUTONUM = @AUTONUM_CS ", parametros, commandTimeout: Config.QueryTimeoutInSeconds());

                    return "YARD deletado com sucesso!";
                }
                catch (Exception erro)
                {
                    return erro.ToString();
                }
            }
        }

        public IEnumerable<Avarias> ListaTipoAvarias()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT CODE AS AUTONUM, DESCR AS DESCRICAO FROM  SGIPA..DTE_TB_AVARIAS WHERE IDENT = 'T' ORDER BY DESCR ");

                var query = con.Query<Avarias>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public IEnumerable<Local> ListaLocalAvarias()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT CODE AS AUTONUM, DESCR AS DESCRICAO FROM  SGIPA..DTE_TB_AVARIAS WHERE IDENT = 'L' ORDER BY DESCR ");

                var query = con.Query<Local>(sb.ToString()).AsEnumerable();

                return query;


            }

        }
        public IEnumerable<Complemento> ListaComplementoAvarias()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT CODE AS AUTONUM, DESCR AS DESCRICAO FROM  SGIPA..DTE_TB_AVARIAS WHERE IDENT = 'C' ORDER BY DESCR");

                var query = con.Query<Complemento>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public IEnumerable<Providencia> ListaProvidenciaAvarias()
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT CODE AS AUTONUM, DESCR AS DESCRICAO FROM  SGIPA..DTE_TB_PROVIDENCIAS ORDER BY DESCR ");

                var query = con.Query<Providencia>(sb.ToString()).AsEnumerable();

                return query;


            }

        }

        public IEnumerable<ConteinersDTO> ConsultarClasseCT(string armazem)

        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                StringBuilder sb = new StringBuilder();


                sb.Clear();
                sb.Append(" SELECT DISTINCT  A.AUTONUM AS AUTONUM_TIPO_CT, A.DESCR AS DESCR_TIPO_CT FROM SGIPA..TB_CAD_TIPO_CT A ");
                sb.Append(" INNER JOIN SGIPA..TB_ARMAZENS_IPA B ON A.AUTONUM = B.TIPO_CT ");
                sb.Append(" WHERE B.DESCR = '" + armazem + "' ");
                sb.Append(" ORDER BY A.DESCR  ");

                var query = con.Query<ConteinersDTO>(sb.ToString()).AsEnumerable();

                return query;


            }
        }

        public LotesDTO ObterDadosHubLote(int lote)
        {
            using (SqlConnection con = new SqlConnection(Config.StringConexao()))
            {
                var parametros = new DynamicParameters();
                parametros.Add("lote", lote, direction: ParameterDirection.Input);

                return con.Query<LotesDTO>(@"
                      SELECT HUB, CIDADE_HUB FROM SGIPA..VW_BL_CARGA_CD WHERE LOTE = @lote ", parametros, commandTimeout: Config.QueryTimeoutInSeconds()).FirstOrDefault();

            }
        }

    }
}