USE [SGIPA]
GO

/****** Object:  View [dbo].[VW_CNTRIPA_BL]    Script Date: 20/05/2020 17:02:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[VW_CNTRIPA_BL]
AS
SELECT        MAX(a.AUTONUM) AS autonumcntr, MAX(ISNULL(h.TRIGRAMA, ' ')) AS pol, MAX(f.NUMERO) AS numero_bl, MAX(ISNULL(c.FANTASIA, ' ')) AS importador, MAX(ISNULL(I.FANTASIA, ' ')) AS INDICADOR, MAX(td.DESCR) 
                         AS tipo_documento, MAX(CASE WHEN ISNULL(c.flag_retirada_vip, 0) = 0 THEN 'N' WHEN ISNULL(c.flag_retirada_vip, 0) = 1 THEN 'Y' END) AS flag_retirada_vip, MAX(CASE WHEN ISNULL(f.canal_siscomex, 9) 
                         = 9 THEN 0 WHEN ISNULL(f.canal_siscomex, 0) = 0 THEN 2 WHEN ISNULL(f.canal_siscomex, 0) = 1 THEN 3 WHEN ISNULL(f.canal_siscomex, 0) = 2 THEN 1 WHEN ISNULL(f.canal_siscomex, 0) = 3 THEN 4 END) AS status, 
                         MAX(f.FLAG_DTAS) AS FLAG_DTAS, SUM(g.bruto) as bruto,MAX(ISNULL(I.FLAG_ALERTA_COL_INDICADOR, 0)) AS FLAG_ALERTA_COL_INDICADOR,
						 MAX(ISNULL(F.NUM_DOCUMENTO,' ')) AS NUM_DOCUMENTO, max(CONVERT(CHAR,F.DT_DESEMBARACO,103)) AS DT_DESEMBARACO
FROM            dbo.TB_CNTR_BL AS a INNER JOIN
                         dbo.TB_AMR_CNTR_BL AS g ON a.AUTONUM = g.CNTR INNER JOIN
                         dbo.TB_BL AS f ON g.BL = f.AUTONUM LEFT OUTER JOIN
                         dbo.TB_TIPOS_DOCUMENTOS AS td ON f.TIPO_DOCUMENTO = td.CODE LEFT OUTER JOIN
                         dbo.TB_CAD_PARCEIROS AS c ON f.IMPORTADOR = c.AUTONUM LEFT OUTER JOIN
                         dbo.TB_CAD_PARCEIROS AS I ON f.CAPTADOR = I.AUTONUM LEFT OUTER JOIN
                         dbo.DTE_TB_PORTOS AS h ON f.PORTO_ORIGEM = h.CODE
WHERE        (f.FLAG_ATIVO = 1) AND (a.FLAG_TERMINAL = 1) AND (a.FLAG_HISTORICO = 0) AND (f.ULTIMA_SAIDA IS NULL)
GROUP BY a.AUTONUM
GO

USE [OPERADOR]
GO


CREATE VIEW [dbo].[VW_INVENT_SISTEMAS_COL] ([id_conteiner], [bruto], [tamanho], [tipobasico], [ef],  [pod], 
[yard], [lacre1], [lacre2], [lacre3], [lacre4], [lacre5], [lacre6], [lacre7], [data_ent_temp], [temperature], [scale], [imo1],[imo2],[imo3],[imo4],
 [impexp], [REFERENCE], [descricao], [fantasia], [nome], [viagem], [stowage], [descr], [STATUS], [autonum], [patio], 
 [viagem_codesp], [flag_csi], [under_deck], [flag_destinado_desova], [flag_bloqueio_mapa], [flag_bloqueio], 
 [flag_bloqueio_manual], [autonumviagem], [sistema], [autonum_via_redex], [joint], [flag_tranship], [flag_reefer], [flag_oog],
 [dlv_term], [doc_type], [cntr_reg], [gwt_code], [flag_spc], [flag_vip], [underdeck], [UMIDADE], [VENTILACAO], [finality], 
 [flag_bloqueio_dte],
 [FLAG_TRS], [STREET_ROW],  [flag_dta], [FLAG_REEFER_DESLIGADO], [INDICADOR]
 ,[FLAG_ALERTA_COL_INDICADOR],[TARA],[TEMP],[TEMPMIN],[TEMPMAX],[ISO],[MAXGROSS],[LACRE_ORIGEM],[LACRE_IPA]
 ,[LACRE1_GATE],[LACRE2_GATE],[LACRE3_GATE],[GMCI],[DESCRICAO_SIGVIG],[NUM_DOCUMENTO],[DT_DESEMBARACO],[FLAG_CNTR_BLOQ_CHAO_PATIO]
 ) 
AS
SELECT     a.id_conteiner, vw.bruto, ISNULL(a.tamanho, 0) AS tamanho, ISNULL(b.codigo, ' ') AS tipobasico, CASE WHEN ISNULL(a.flag_desovado, 0) 
                      = 0 THEN 'F' WHEN ISNULL(a.flag_desovado, 0) = 1 THEN 'E' END AS ef, 'SSZ' AS pod, ISNULL(a.yard, ' ') AS yard, isnull(a.lacre_origem,' ') AS lacre1,
                       ISNULL(a.lacre2, ' ') AS lacre2, ISNULL(a.lacre3, ' ') AS lacre3, ISNULL(a.lacre4, ' ') AS lacre4, ' ' AS lacre5, ' ' AS lacre6, ' ' AS lacre7, CONVERT(VARCHAR, 
                      a.dt_entrada, 103) AS data_ent_temp, /* 'DD/MM/YYYY') AS data_ent_temp,*/ ISNULL(a.temp, ' ') AS temperature, ' ' AS scale, ISNULL(a.imo1, ' ') AS imo1, ISNULL(a.imo2, ' ') AS imo2,ISNULL(a.imo3, ' ') AS imo3,ISNULL(a.imo4, ' ') AS imo4,   
                      'I' AS impexp, vw2.numero_bl AS REFERENCE, a.regime AS descricao, ISNULL(vw2.importador, ' ') AS fantasia, E.DESCR AS nome, '.' AS viagem, ' ' AS stowage, 
                      ' ' AS descr, ISNULL(vw2.status, 0) AS status, a.autonum, CASE A.PATIO WHEN '7' THEN 1 ELSE CONVERT(INT, a.patio) END AS patio, a.viagem AS viagem_codesp, 0 AS flag_csi, '-' AS under_deck, 
                      ISNULL(a.flag_destinado_desova, 0) AS flag_destinado_desova, 0 AS flag_bloqueio_mapa, ISNULL(a.flag_bloqueio, 0) AS flag_bloqueio, 
                      ISNULL(a.flag_bloqueio_manual, 0) AS flag_bloqueio_manual, 0 AS autonumviagem, 'I' AS sistema, 0 AS autonum_via_redex, 0 AS joint, 'N' as flag_tranship, 
                      CASE WHEN ISNULL(a.temp, ' ') = ' ' THEN 'N' WHEN ISNULL(a.temp, ' ') <> ' ' THEN 'Y' END AS flag_reefer, CASE WHEN ISNULL(a.overheight, 0) 
                      + ISNULL(a.overwidth, 0) + ISNULL(a.overwidth_r, 0) + ISNULL(a.overlength, 0) = 0 THEN 'N' WHEN ISNULL(a.overheight, 0) + ISNULL(a.overwidth, 0) 
                      + ISNULL(a.overwidth_r, 0) + ISNULL(a.overlength, 0) > 0 THEN 'Y' END AS flag_oog, '' AS dlv_term, ISNULL(vw2.tipo_documento, ' ') AS doc_type, 
                      a.regime AS cntr_reg, CASE WHEN vw.bruto <= 12000 THEN 'I' WHEN vw.bruto > 12000 AND vw.bruto <= 20000 THEN 'II' WHEN vw.bruto > 20000 AND 
                      vw.bruto <= 24000 THEN 'III' WHEN vw.bruto > 24000 THEN 'IV' END AS gwt_code, 'N' AS flag_spc, vw2.flag_retirada_vip as flag_vip, 0 AS underdeck, isnull(a.umidade,' ') AS UMIDADE, 
                      isnull(a.ventilacao,' ') AS VENTILACAO, 'IPA-IMPORT.' as finality,  
						0 AS FLAG_BLOQUEIO_DTE,  
                      0 AS FLAG_TRS,ISNULL(a.STREET_ROW, ' ') AS STREET_ROW,
                      CASE WHEN ISNULL(vw2.FLAG_DTAS, 0) = 0 THEN 'N' WHEN ISNULL(vw2.FLAG_DTAS, 0) = 1 THEN 'Y' END AS flag_dta,
					  CASE ISNULL(A.FLAG_DESLIGADO,0) WHEN 0 THEN 'N' WHEN 1 THEN 'Y' END AS FLAG_REEFER_DESLIGADO,
					  vw2.INDICADOR
					  ,ISNULL(VW2.FLAG_ALERTA_COL_INDICADOR,0) AS FLAG_ALERTA_COL_INDICADOR ,A.TARA,ISNULL(A.TEMP, ' ') AS TEMP, ISNULL(A.TEMPMIN,' ') AS TEMPMIN, ISNULL(A.TEMPMAX,' ') AS TEMPMAX ,ISNULL(A.ISO,' ') AS ISO,ISNULL(A.MAXGROSS,0) ,ISNULL(A.LACRE_ORIGEM,' ') AS LACRE_ORIGEM
					  ,ISNULL(A.LACRE_IPA,' ') AS LACRE_IPA,ISNULL(A.LACRE1_GATE,' ') AS LACRE1_GATE,ISNULL(A.LACRE2_GATE,' ') AS LACRE2_GATE,ISNULL(A.LACRE3_GATE,' ') AS LACRE3_GATE ,ISNULL(A.GMCI,' ') AS GMCI , ISNULL(A.DESCRICAO_SIGVIG, ' ') AS DESCRICAO_SIGVIG,
					  ISNULL(VW2.NUM_DOCUMENTO, ' ') AS NUM_DOCUMENTO, ISNULL(VW2.DT_DESEMBARACO,' '),ISNULL(A.FLAG_CNTR_BLOQ_CHAO_PATIO,0)
					  FROM         SGIPA..tb_cntr_bl a INNER JOIN
                      SGIPA..vw_bruto_cntr vw ON a.autonum = vw.autonum LEFT JOIN
                      SGIPA..dte_tb_viagens d ON a.viagem = d .viagem LEFT JOIN
                      SGIPA..dte_tb_navios e ON d .navio = e.code LEFT JOIN
                      SGIPA..dte_tb_tipos_conteiner b ON a.tipo = b.code LEFT JOIN
                      SGIPA..vw_cntripa_bl vw2 ON a.autonum = vw2.autonumcntr
           WHERE     (a.flag_historico = 0) AND (a.flag_terminal = 1) AND ISNULL(a.autonumviagem, 0) = 0
 UNION ALL
SELECT     a.DESCR, 0 AS bruto, ISNULL(a.tamanho_ct, 0) AS tamanho, ISNULL(a.tipobasico_ct, ' ') AS tipobasico, 'F' AS ef,  'SSZ' AS pod,  
                      ISNULL(a.yard, ' ') AS yard, ' ' AS lacre1, ' ' AS lacre2, ' ' AS lacre3, ' ' AS lacre4, ' ' AS lacre5, ' ' AS lacre6, ' ' AS lacre7, CONVERT(VARCHAR, a.dt_entrada, 103), 
                      ' ' AS temperature, ' ' AS escala, ' ' AS imo1,' ' AS imo2,' ' AS imo3, ' ' AS imo4,'Z' AS impexp, ' ' AS REFERENCE, ' ' AS descricao, ' ' AS fantasia, ' ' AS nome, ' ' AS viagem, ' ' AS stowage, ' ' AS descr, 
                      0 AS status, a.AUTONUM AS autonum, CASE A.PATIO WHEN 7 THEN 1 ELSE ISNULL(CONVERT(INT, a.patio), 0) END AS patio, ' ' AS viagem_codesp, 0 AS flag_csi, '-' AS under_deck, 0 AS flag_destinado_desova,
                       0 AS flag_bloqueio_mapa, 0 AS flag_bloqueio, 0 AS flag_bloqueio_manual, 0 AS autonumviagem, 'A' AS sistema, 0, 0 AS joint, 'N', 'N', 'N', ' ' AS dlv_term, 
                      ' ' AS doc_type, ' ' AS cntr_reg, ' ' AS gwt_code, 'N' AS flag_spc, 'N' AS flag_vip, 0 AS underdeck, 'OFF' AS UMIDADE, 'CLOSED' AS VENTILACAO, 'WH. CNTR', 
                          0, 0 AS FLAG_TRS
,ISNULL(a.STREET_ROW, ' ') AS STREET_ROW,'N' AS FLAG_DTA, 'N' AS FLAG_REEFER_DESLIGADO,' ' AS INDICADOR
,0,0
,' ',' ',' ',' ',0,' '
,' ',' ',' ',' ',' ',' ' ,' ',' ',0
FROM         SGIPA..tb_ARMAZENS_ipa a
WHERE     A.FLAG_CT = 1 AND A.DT_SAIDA IS NULL
GO



CREATE FUNCTION fnInvent_Sistemas_Yard (@Sigla varchar(12))
returns @rtTable table
(
    id_conteiner nvarchar(12),
	tamanho nvarchar(2),
	tipobasico nvarchar(2),
	ef nvarchar(2),
	yard nvarchar(15),
	autonum numeric(8),
	patio numeric(1),
	sistema nvarchar(1)
)
as
begin
    insert into @rtTable
    select top (1) id_conteiner,tamanho,tipobasico,ef,yard,autonum,patio,sistema
    from operador.dbo.VW_INVENT_SISTEMAS_YARD
    where id_conteiner=@Sigla
   return;
end
GO

CREATE FUNCTION fnInvent_Sistemas_YardF (@FinalSigla varchar(12),@Patio numeric(1))
returns @rtTable table
(
    id_conteiner nvarchar(12),
	tamanho nvarchar(2),
	tipobasico nvarchar(2),
	ef nvarchar(2),
	yard nvarchar(15),
	autonum numeric(8),
	patio numeric(1),
	sistema nvarchar(1)
)
as
begin
	if @Patio=0 
		begin
			insert into @rtTable
			select id_conteiner,tamanho,tipobasico,ef,yard,autonum,patio,sistema
			from operador.dbo.VW_INVENT_SISTEMAS_YARD
			where substring(id_conteiner,8,5)=@FinalSigla
		end
	if @Patio=1 or @Patio=7
		begin
			insert into @rtTable
			select id_conteiner,tamanho,tipobasico,ef,yard,autonum,patio,sistema
			from operador.dbo.VW_INVENT_SISTEMAS_YARD
			where substring(id_conteiner,8,5)=@FinalSigla
			and patio in (1,7)
		end
	if @Patio<>0 and @Patio<>1 and @Patio<>7
		begin
			insert into @rtTable
			select id_conteiner,tamanho,tipobasico,ef,yard,autonum,patio,sistema
			from operador.dbo.VW_INVENT_SISTEMAS_YARD
			where substring(id_conteiner,8,5)=@FinalSigla
			and patio=@Patio
		end
   return;
end
GO

CREATE FUNCTION fnInvent_Sistemas_Col (@Sigla varchar(12),@Patio numeric(1))
returns @rtTable table
(
    autonum numeric(8),
	id_conteiner nvarchar(12),
	patio numeric(1),
	ef nvarchar(1),
	nome nvarchar(50),
	viagem nvarchar(20),
	tamanho numeric(2),
	tipobasico nvarchar(2),
	fantasia nvarchar(75),
	data_ent_temp nvarchar(30),
	yard nvarchar(15),
	pod nvarchar(15),
	imo1 nvarchar(8),
	lacre1 nvarchar(20),
	lacre2 nvarchar(20),
	lacre3 nvarchar(20),
	lacre4 nvarchar(20),
	lacre5 nvarchar(20),
	lacre6 nvarchar(20),
	lacre7 nvarchar(20),
	bruto numeric(8),
	sistema nvarchar(1),
	flag_reefer nvarchar(1),
	flag_reefer_desligado nvarchar(1),
	flag_dta nvarchar(1),
	flag_destinado_desova numeric(1),
	indicador nvarchar(75),
	flag_alerta_col_indicador numeric(1),
	tara numeric(8),
	temp nvarchar(8),
	tempmin nvarchar(8),
	tempmax nvarchar(8),
	iso nvarchar(4),
	maxgross numeric(6),
	lacre_origem nvarchar(20),
	lacre_ipa nvarchar(20),
	lacre1_gate nvarchar(20),
	lacre2_gate nvarchar(20),
	lacre3_gate nvarchar(20),
	gmci nvarchar(20),
	descricao_sigvig nvarchar(200),
	num_documento nvarchar(25),
	flag_bloqueio_manual numeric(1),
	flag_bloqueio numeric(1),
	dt_desembaraco nvarchar(20),
	FLAG_CNTR_BLOQ_CHAO_PATIO NUMERIC(1)
)
as
begin
	if @Patio<>1 and @Patio<>7
	begin
		insert into @rtTable
		select autonum,id_conteiner,patio,ef,nome,viagem,tamanho,
				tipobasico,fantasia,data_ent_temp,yard,pod,imo1,
				lacre1,lacre2,lacre3,lacre4,lacre5,lacre6,lacre7,
				bruto,sistema,flag_reefer,flag_reefer_desligado,
				flag_dta,flag_destinado_desova,indicador,
				flag_alerta_col_indicador,tara,temp,tempmin,
				tempmax,iso,maxgross,lacre_origem,lacre_ipa,
				lacre1_gate,lacre2_gate,lacre3_gate,gmci,
				descricao_sigvig,num_documento,flag_bloqueio_manual,
				flag_bloqueio,dt_desembaraco,FLAG_CNTR_BLOQ_CHAO_PATIO
		from operador.dbo.VW_INVENT_SISTEMAS_COL
		where  id_conteiner=@Sigla and patio=@Patio
	end
	if @Patio=1 or @Patio=7
	begin
		insert into @rtTable
		select autonum,id_conteiner,patio,ef,nome,viagem,tamanho,
				tipobasico,fantasia,data_ent_temp,yard,pod,imo1,
				lacre1,lacre2,lacre3,lacre4,lacre5,lacre6,lacre7,
				bruto,sistema,flag_reefer,flag_reefer_desligado,
				flag_dta,flag_destinado_desova,indicador,
				flag_alerta_col_indicador,tara,temp,tempmin,
				tempmax,iso,maxgross,lacre_origem,lacre_ipa,
				lacre1_gate,lacre2_gate,lacre3_gate,gmci,
				descricao_sigvig,num_documento,flag_bloqueio_manual,
				flag_bloqueio,dt_desembaraco,FLAG_CNTR_BLOQ_CHAO_PATIO
		from operador.dbo.VW_INVENT_SISTEMAS_COL
		where id_conteiner=@Sigla and patio in(1,7)
	end
   return;
end
GO

CREATE INDEX IDX_HIST_CNTR_001 ON TB_HIST_SHIFTING(CNTR,TIPO,ORIGEM)
GO


ALTER FUNCTION [dbo].[fnMovimentaCntr] (@Sigla nvarchar(12),@Patio numeric(1),@Tamanho numeric(2),@Sistema nvarchar(1),
@YardOrigem nvarchar(10),@YardDestino nvarchar(10),@Usuario numeric(8),@FLAG_CNTR_BLOQ_CHAO_PATIO NUMERIC(1)
)
returns varchar(300)
as
begin
	DECLARE @Mensagem nvarchar(300);
	DECLARE @ID INT;
	DECLARE @FLAG_BLOQUEIO_POS NUMERIC(1);
	DECLARE @VALIDAORIGEM NUMERIC(1);
	DECLARE @VALIDADESTINO NUMERIC(1);
    DECLARE @ColQuadraD NVARCHAR(2);
    DECLARE @ColFiadaD NVARCHAR(2);
    DECLARE @ColLastroD NVARCHAR(2);
    DECLARE @ColAlturaD NVARCHAR(1);
    DECLARE @ColPilhaD NVARCHAR(2);
    DECLARE @ColQuadraO NVARCHAR(2);
    DECLARE @ColFiadaO NVARCHAR(2);
    DECLARE @ColLastroO NVARCHAR(2);
    DECLARE @ColAlturaO NVARCHAR(1);
    DECLARE @ColPilhaO NVARCHAR(2);
	DECLARE @YardAbaixo NVARCHAR(7);
	DECLARE @YardAcima NVARCHAR(7);
	DECLARE @YardPos NVARCHAR(7);
	DECLARE @YardAnt NVARCHAR(7);
	DECLARE @SiglaConc NVARCHAR(12);
	DECLARE @SiglaAcima NVARCHAR(12);
	

	SET @Mensagem='OK';
	IF @Mensagem='OK'
	BEGIN
		SELECT @ID=ISNULL(MAX(AUTONUM),0),@FLAG_BLOQUEIO_POS=ISNULL(MAX(FLAG_BLOQUEIO),0),@VALIDADESTINO=ISNULL(MAX(VALIDA),0) FROM TB_YARD WHERE PATIO=@Patio AND YARD=@YardDestino;
		SELECT @VALIDAORIGEM=ISNULL(MAX(VALIDA),0) FROM TB_YARD WHERE PATIO=@Patio AND YARD=@YardOrigem;
		IF @ID=0 
		BEGIN
			SET @Mensagem='Posicao nao encontrada no patio';
		END
		IF @Mensagem='OK'
		BEGIN
			IF @FLAG_BLOQUEIO_POS=1
			BEGIN
				SET @Mensagem='Posicao Bloqueada no patio';
			END
		END
		IF @Mensagem='OK'
		BEGIN
			IF @PATIO=3 AND @VALIDADESTINO=1
			BEGIN
				SET @ColQuadraD=SUBSTRING(@YardDestino,1,2);
				SET @ColFiadaD=SUBSTRING(@YardDestino,3,2);
				SET @ColLastroD=SUBSTRING(@YardDestino,5,2);
				SET @ColAlturaD=SUBSTRING(@YardDestino,7,1);
				SET @ColPilhaD = @ColQuadraD + @ColFiadaD + @ColLastroD
			END
			IF @PATIO<>3 AND @VALIDADESTINO=1
			BEGIN
				SET @ColQuadraD=SUBSTRING(@YardDestino,1,1);
				SET @ColFiadaD=SUBSTRING(@YardDestino,2,2);
				SET @ColLastroD=SUBSTRING(@YardDestino,4,2);
				SET @ColAlturaD=SUBSTRING(@YardDestino,6,1);
				SET @ColPilhaD = @ColQuadraD + @ColFiadaD + @ColLastroD;
			END
			IF @PATIO=3 AND @VALIDAORIGEM=1
			BEGIN
				SET @ColQuadraO=SUBSTRING(@YardOrigem,1,2);
				SET @ColFiadaO=SUBSTRING(@YardOrigem,3,2);
				SET @ColLastroO=SUBSTRING(@YardOrigem,5,2);
				SET @ColAlturaO=SUBSTRING(@YardOrigem,7,1);
				SET @ColPilhaO = @ColQuadraO + @ColFiadaO + @ColLastroO
			END
			IF @PATIO<>3 AND @VALIDAORIGEM=1
			BEGIN
				SET @ColQuadraO=SUBSTRING(@YardOrigem,1,1);
				SET @ColFiadaO=SUBSTRING(@YardOrigem,2,2);
				SET @ColLastroO=SUBSTRING(@YardOrigem,4,2);
				SET @ColAlturaO=SUBSTRING(@YardOrigem,6,1);
				SET @ColPilhaO = @ColQuadraO + @ColFiadaO + @ColLastroO;
			END
			IF @MENSAGEM='OK' AND @VALIDADESTINO=1
			BEGIN
				IF @ColAlturaD='1' AND @FLAG_CNTR_BLOQ_CHAO_PATIO=1
				BEGIN
					SET @MENSAGEM='Bloqueio de CHAO cadastrado para o conteiner'
				END
			END
			
			--SE MOVER DE PILHA NAO PODE TER UM ACIMA NA ORGIEM
			IF @MENSAGEM='OK' AND @VALIDAORIGEM=1
			BEGIN
				SET @YardAcima=@ColQuadraO + @ColFiadaO + @ColLastroO + dbo.udf_CIntToChar(CONVERT(INT,@ColAlturaO) + 1,1)
				SELECT @ID=ISNULL(MAX(AUTONUM),0),@SiglaAcima=ISNULL(max(ID_CONTEINER),'') FROM VW_INVENT_SISTEMAS_YARD WHERE PATIO=@PATIO AND YARD=@YardAcima
				IF @ID<>0 
				BEGIN
					SET @MENSAGEM='Existe um Conteiner acima (' + @SiglaAcima + '), impossivel movimentar  ' 
				END
			END
			IF @MENSAGEM='OK' AND @VALIDADESTINO=1 
			BEGIN
				IF @Tamanho=20 AND dbo.udf_CIntToChar(@ColFiadaD,1) % 2=0 
				BEGIN
					SET @MENSAGEM='Conteiner de 20 deve ter fiada impar'
				END 
				IF @Tamanho=40 AND dbo.udf_CIntToChar(@ColFiadaD,1) % 2=1 
				BEGIN
					SET @MENSAGEM='Conteiner de 40 deve ter fiada par'
				END 
				--ALTURA ANTERIOR
				IF @MENSAGEM='OK' AND @ColAlturaD<>'1'
				BEGIN
					SET @YardAbaixo=@ColQuadraD + @ColFiadaD + @ColLastroD + dbo.udf_CIntToChar(CONVERT(INT,@ColAlturaD) - 1,1)
					SELECT @ID=ISNULL(MAX(AUTONUM),0) FROM VW_INVENT_SISTEMAS_YARD WHERE PATIO=@PATIO AND YARD=@YardAbaixo
					IF @ID=0 
					BEGIN
						SET @MENSAGEM='Falta Conteiner na altura anterior'
					END
				END
				--CONCOMITANCIA
				IF @MENSAGEM='OK' AND @TAMANHO=20 AND @VALIDADESTINO=1
				BEGIN
					SET @YardAnt=@ColQuadraD + dbo.udf_CIntToChar(CONVERT(INT,@ColFiadaD) - 1,2)  + @ColLastroD + @ColAlturaD
					SET @YardPos=@ColQuadraD + dbo.udf_CIntToChar(CONVERT(INT,@ColFiadaD) + 1,2)  + @ColLastroD + @ColAlturaD
					SELECT @ID=ISNULL(MAX(AUTONUM),0),@SiglaConc=ISNULL(MAX(ID_CONTEINER),'') FROM VW_INVENT_SISTEMAS_YARD WHERE PATIO=@PATIO AND YARD IN(@YardDestino,@YardAnt,@YardPos) AND ID_CONTEINER<>@SIGLA
					IF @ID>0
					BEGIN
						SET @MENSAGEM='Posicao já ocupada pelo conteiner ' + @SiglaConc
					END
				END
				IF @MENSAGEM='OK' AND @TAMANHO=40 AND @VALIDADESTINO=1
				BEGIN
					SET @YardAnt=@ColQuadraD + dbo.udf_CIntToChar(CONVERT(INT,@ColFiadaD) - 1,2)  + @ColLastroD + @ColAlturaD
					SET @YardPos=@ColQuadraD + dbo.udf_CIntToChar(CONVERT(INT,@ColFiadaD) + 1,2)  + @ColLastroD + @ColAlturaD
					SELECT @ID=ISNULL(MAX(AUTONUM),0),@SiglaConc=ISNULL(MAX(ID_CONTEINER),'') FROM VW_INVENT_SISTEMAS_YARD WHERE PATIO=@PATIO AND YARD IN(@YardDestino,@YardAnt,@YardPos) AND ID_CONTEINER<>@SIGLA
					IF @ID>0
					BEGIN
						SET @MENSAGEM='Posicao já ocupada pelo conteiner ' + @SiglaConc
					END
					IF @MENSAGEM='OK'
					BEGIN
						SET @YardAnt=@ColQuadraD + dbo.udf_CIntToChar(CONVERT(INT,@ColFiadaD) - 2,2)  + @ColLastroD + @ColAlturaD
						SET @YardPos=@ColQuadraD + dbo.udf_CIntToChar(CONVERT(INT,@ColFiadaD) + 2,2)  + @ColLastroD + @ColAlturaD
						SELECT @ID=ISNULL(MAX(AUTONUM),0),@SiglaConc=ISNULL(MAX(ID_CONTEINER),'') FROM VW_INVENT_SISTEMAS_YARD WHERE PATIO=@PATIO AND YARD IN(@YardDestino,@YardAnt,@YardPos) AND ID_CONTEINER<>@SIGLA
						IF @ID>0
						BEGIN
							SET @MENSAGEM='Posicao já ocupada pelo conteiner ' + @SiglaConc
						END
					END
				END
			END
			IF @MENSAGEM='OK'
			BEGIN
				SELECT @MENSAGEM = ISNULL(max(B1.OBS),'OK') FROM INTELOPER..TB_IP_ATRIBUTO_YARD A1
				Inner Join INTELOPER..TB_IP_ATRIBUTOS B1 ON A1.AUTONUM_ATR = B1.AUTONUM_ATR,
				(SELECT MAX (A.PATIO) AS PATIO,MAX (A.YARD) AS YARD,MAX (A.AUTONUM_ATR) AS AUTONUM_ATR,
				MAX (B.DESCR_ATR) AS DESCR_ATR,MAX (A.STATUS_ATRIB) AS STATUS_ATRIB,MAX (a.DT_INI_ATR) As DT_INI_ATR
				FROM INTELOPER..TB_IP_ATRIBUTO_YARD A Inner Join INTELOPER..TB_IP_ATRIBUTOS b ON A.AUTONUM_ATR = B.AUTONUM_ATR
				Where a.DT_INI_ATR <= GETDATE() AND (A.DT_FIM_ATR >= GETDATE() OR A.DT_FIM_ATR IS NULL)
				GROUP BY A.PATIO,
				A.YARD,A.AUTONUM_ATR,B.DESCR_ATR,A.STATUS_ATRIB) C1,
				(SELECT * FROM OPERADOR..VW_INVENT_SISTEMAS_COL WHERE ID_CONTEINER=@Sigla) I Where A1.Patio = C1.Patio
				AND A1.YARD = C1.YARD AND A1.AUTONUM_ATR = C1.AUTONUM_ATR AND A1.STATUS_ATRIB = C1.STATUS_ATRIB
				AND A1.DT_INI_ATR = C1.DT_INI_ATR AND A1.PATIO=I.PATIO AND (A1.YARD = @YardDestino OR A1.YARD=@ColQuadraD)
				AND ( (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 1 AND A1.VLR_ATRIB = I.VIAGEM) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 1 AND A1.VLR_ATRIB <> I.VIAGEM)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 4 AND A1.VLR_ATRIB = I.TAMANHO) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 4 AND A1.VLR_ATRIB <> I.TAMANHO)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 5 AND A1.VLR_ATRIB = I.TIPOBASICO) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 5 AND A1.VLR_ATRIB <> I.TIPOBASICO)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 6 AND A1.VLR_ATRIB = I.POD) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 6 AND A1.VLR_ATRIB <> I.POD)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 7 AND A1.VLR_ATRIB = I.EF) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 7 AND A1.VLR_ATRIB <> I.EF)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 10 AND A1.VLR_ATRIB = I.DOC_TYPE) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 10 AND A1.VLR_ATRIB <> I.DOC_TYPE)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 11 AND A1.VLR_ATRIB = I.CNTR_REG) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 11 AND A1.VLR_ATRIB <> I.CNTR_REG)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 12 AND A1.VLR_ATRIB = I.IMO1) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 12 AND A1.VLR_ATRIB <> I.IMO1)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 13 AND A1.VLR_ATRIB = CASE I.FLAG_REEFER WHEN 'Y' THEN 1 ELSE 0 END ) 
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 13 AND A1.VLR_ATRIB <> CASE I.FLAG_REEFER WHEN 'Y' THEN 1 ELSE 0 END )
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 14 AND A1.VLR_ATRIB = CASE I.FLAG_OOG WHEN 'Y' THEN 1 ELSE 0 END )
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 14 AND A1.VLR_ATRIB <> CASE I.FLAG_OOG WHEN 'Y' THEN 1 ELSE 0 END )
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 17 AND A1.VLR_ATRIB = I.GWT_CODE)
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 17 AND A1.VLR_ATRIB <> I.GWT_CODE)
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 19 AND A1.VLR_ATRIB = I.FINALITY)
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 19 AND A1.VLR_ATRIB <> I.FINALITY) 
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB =  CASE  RTRIM(LTRIM(I.IMO1)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB <> CASE   RTRIM(LTRIM(I.IMO1)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB =  CASE  RTRIM(LTRIM(I.IMO2)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB <> CASE   RTRIM(LTRIM(I.IMO2)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB =  CASE  RTRIM(LTRIM(I.IMO3)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB <> CASE   RTRIM(LTRIM(I.IMO3)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB =  CASE  RTRIM(LTRIM(I.IMO4)) WHEN '' THEN 0 ELSE 1 END ) 
				OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB <> CASE   RTRIM(LTRIM(I.IMO4)) WHEN '' THEN 0 ELSE 1 END ) 
				)
			END
		END
	END
	
	
	
	return @Mensagem;
end
GO




CREATE Procedure pcMovimentaCntr(@Autonum numeric(8),@Sigla nvarchar(12),@Gmci nvarchar(15),@Sistema nvarchar(1)
,@YardOrigem nvarchar(10),@YardDestino nvarchar(10),@Tara numeric(5), @Lacre_ipa nvarchar(20), @Lacre1_Gate nvarchar(20)
,@Lacre2_Gate nvarchar(20), @Lacre3_Gate nvarchar(20),@Camera_Posicionamento numeric(5),@Usuario numeric(8),@Motivo numeric(8)
) 
WITH 
EXECUTE AS CALLER
AS
BEGIN
	if @Sistema='I'
	BEGIN
		 UPDATE sgipa..tb_cntr_bl SET YARD=@YardDestino,TARA=@Tara,TARA_REAL_EST='R'
		,LACRE_IPA=@LACRE_IPA,LACRE1_GATE=@LACRE1_GATE
		,LACRE2_GATE=@LACRE2_GATE,LACRE3_GATE=@LACRE3_GATE
		,CAMERA_POSICIONAMENTO=@CAMERA_POSICIONAMENTO
		WHERE AUTONUM=@Autonum;

		UPDATE OPERADOR..tb_GATE_DADOS SET LACREIPA=@LACRE_IPA,LACRE1=@LACRE1_GATE
		,LACRE2=@LACRE2_GATE,LACRE3=@LACRE3_GATE
		WHERE CONTEINER=@Sigla
		 AND GMCI=@GMCI;

	END
	if @Sistema='A' 
	BEGIN
		UPDATE sgipa..tb_armazens_ipa SET YARD=@YardDestino
		WHERE AUTONUM=@Autonum;
	END
	if @Sistema='R' 
	BEGIN
		UPDATE REDEX..tb_PATIO SET YARD=@YardDestino
		WHERE AUTONUM_PATIO=@Autonum;
	END
	INSERT INTO OPERADOR..TB_HIST_SHIFTING (CNTR,ORIGEM,DESTINO,DATA,TIPO,USUARIO,MOTIVO) VALUES 
	(@Autonum,@YardOrigem,@YardDestino,Getdate(),@Sistema,@Usuario,@Motivo);
END 



use [SGIPA]
GO

CREATE FUNCTION fnInvent_Sistemas_Arm_Col (@Marcante numeric(12))
returns @rtTable table
(AUTONUM numeric(10),
 AUTONUM_CARGA numeric(10),
 AUTONUM_CARGA_AC numeric(10),
 VOLUMES numeric(10),
 AUTONUM_CS_YARD numeric(10),
 AUTONUM_CS_YARD_AC numeric(10),
 FLAG_RECONHECIMENTO numeric(1),
 LOTE_IMPRESSO numeric(10),
 ID_CONTEINER_IMPRESSO nvarchar(12),
 ARMAZEM_REC numeric(10),
 YARD_REC nvarchar(12),
 AUTONUM_CEXP numeric(10)
)
as
begin
	insert into @rtTable
	SELECT AUTONUM, ISNULL(AUTONUM_CARGA,0) AS AUTONUM_CARGA, ISNULL(AUTONUM_CARGA_AC,0) AS AUTONUM_CARGA_AC ,VOLUMES,
	ISNULL(AUTONUM_CS_YARD,0) AS AUTONUM_CS_YARD,ISNULL(AUTONUM_CS_YARD_AC,0) AS AUTONUM_CS_YARD_AC, 
	ISNULL(FLAG_RECONHECIMENTO,0) AS FLAG_RECONHECIMENTO, ISNULL(LOTE_IMPRESSO,0) AS LOTE_IMPRESSO, 
	ISNULL(ID_CONTEINER_IMPRESSO,' ') AS ID_CONTEINER_IMPRESSO,ISNULL(ARMAZEM_REC,0) AS ARMAZEM_REC, ISNULL(YARD_REC,' ') AS YARD_REC,
	ISNULL(AUTONUM_CEXP,0) AS AUTONUM_CEXP 
	FROM SGIPA..TB_MARCANTES WHERE AUTONUM=@Marcante 
	AND (ISNULL(AUTONUM_CARGA,0)>0 or ISNULL(AUTONUM_CARGA_AC,0)>0 OR ISNULL(FLAG_RECONHECIMENTO,0)=1  or ISNULL(AUTONUM_CEXP,0)>0)
   return;
end

GO


CREATE FUNCTION fnInvent_Sistemas_Arm_ColLote (@Lote numeric(10), @Marcante numeric(12))
returns @rtTable table
(MERCADORIA nvarchar(300),
 MARCA nvarchar(300),
 ENTRADA nvarchar(25),
 IMPORTADOR NVARCHAR(100),
 CNTR_DESOVA NVARCHAR(12),
 TIPO_DOC NVARCHAR(30),
 DESCR_ARMAZEM NVARCHAR(30),
 POSICAO NVARCHAR(30),
 CANAL_ALF NVARCHAR(30),
 MOTIVO_PROX_MVTO NVARCHAR(300),
 VOLUME NUMERIC(12,5),
 IMO nvarchar(100),
 NVOCC nvarchar(100),
 BL nvarchar(50)
)
as
begin
	IF @MARCANTE>0
	BEGIN
		insert into @rtTable
		SELECT MERCADORIA,MARCA,CONVERT(CHAR,DATA_ENTRADA,0) AS ENTRADA,IMPORTADOR,CNTR_DESOVA,
		TIPO_DOC,DESCR_ARMAZEM, POSICAO,CASE CANAL_ALF WHEN 0 THEN 'AMARELO' WHEN 1 THEN 'VERMELHO' WHEN 2 THEN 'VERDE' WHEN 3 THEN 'CINZA' WHEN 9 THEN '' END AS CANAL_ALF,
		MOTIVO_PROX_MVTO + ' ' + CONVERT(CHAR,HORAS_PROX_MVTO,0) AS MOTIVO_PROX_MVTO ,VOLUME, IMO, NVOCC, BL
		FROM OPERADOR..VW_INVENT_ARMAZEM
		WHERE LOTE=@Lote and MARCANTE=@MARCANTE
	END
	IF @MARCANTE=0
	BEGIN
		insert into @rtTable
		SELECT MERCADORIA,MARCA,CONVERT(CHAR,DATA_ENTRADA,0) AS ENTRADA,IMPORTADOR,CNTR_DESOVA,
		TIPO_DOC,DESCR_ARMAZEM, POSICAO,CASE CANAL_ALF WHEN 0 THEN 'AMARELO' WHEN 1 THEN 'VERMELHO' WHEN 2 THEN 'VERDE' WHEN 3 THEN 'CINZA' WHEN 9 THEN '' END AS CANAL_ALF,
		MOTIVO_PROX_MVTO + ' ' + CONVERT(CHAR,HORAS_PROX_MVTO,0) AS MOTIVO_PROX_MVTO ,VOLUME, IMO, NVOCC, BL
		FROM OPERADOR..VW_INVENT_ARMAZEM
		WHERE LOTE=@Lote 
	END

   return;
end

GO

CREATE FUNCTION fnInvent_Sistemas_Arm_ColItem (@Lote numeric(10), @Marcante numeric(12), @Autonum_Cs_Yard numeric(12))
returns @rtTable table
(AUTONUM nvarchar(30),
 DISPLAY NVARCHAR(300)
)
as
begin
	IF @MARCANTE>0
	BEGIN
		insert into @rtTable
		select ID_GRAVACAO as autonum, ( REPLACE(rtrim(convert(char,QTDE)),'.000','') +  '/' + rtrim(convert(char,QTDE_CAPTADA)) + ' ' + ISNULL(EMBALAGEM, ' ') + ' ' +  ISNULL(DESCR_ARMAZEM,' ') + ' ' + ISNULL(POSICAO, ' ')  ) AS DISPLAY
		FROM OPERADOR..VW_INVENT_ARMAZEM
		Where
		lote = @Lote
		and Marcante = @Marcante
		and Autonum_CS_Yard =@Autonum_CS_Yard
		AND QTDE>0 ORDER BY DESCR_ARMAZEM,POSICAO
	END
	IF @MARCANTE=0
	BEGIN
		insert into @rtTable
		select ID_GRAVACAO as autonum, ( REPLACE(rtrim(convert(char,QTDE)),'.000','') +  '/' + rtrim(convert(char,QTDE_CAPTADA)) + ' ' + ISNULL(EMBALAGEM, ' ') + ' ' +  ISNULL(DESCR_ARMAZEM,' ') + ' ' + ISNULL(POSICAO, ' ')  ) AS DISPLAY
		FROM OPERADOR..VW_INVENT_ARMAZEM
		Where
		lote = @Lote
		and Autonum_CS_Yard =@Autonum_CS_Yard
		AND QTDE>0 ORDER BY DESCR_ARMAZEM,POSICAO
	END
   return;
end
GO


ALTER Procedure pcMovimentaCS(@Tag nvarchar(1),@Tipo nvarchar(1),@FlagChronos numeric(1),@AutonumCs numeric(10)
,@Patio numeric(2),@Yard nvarchar(10), @Quantidade numeric(8), @Motivo numeric(8)
,@Marcante numeric(12), @Usuario numeric(8)
) 
WITH 
EXECUTE AS CALLER
AS
BEGIN
	Declare @Autonum_Cs_Yard numeric(10)
	
	if @FlagChronos=1
	BEGIN
		INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD(AUTONUM_CS,PATIO,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD)
		VALUES (@AutonumCS,@Patio,0,@Yard,'I',@Quantidade,@Motivo,0,0,0,0,0);
		SELECT @Autonum_Cs_Yard = IDENT_CURRENT('TB_CARGA_SOLTA_YARD');
	END
	IF @FlagChronos=0
	BEGIN
		INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD_AC(AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD)
		VALUES (@AutonumCS,0,@Yard,'I',@Quantidade,@Motivo,0,0,0,0,0);
		SELECT @Autonum_Cs_Yard = IDENT_CURRENT('TB_CARGA_SOLTA_YARD_AC');
	END 
	IF @Tag<>'M' AND @Marcante>'0'
	BEGIN
		IF @FlagChronos=1
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CARGA=@AUTONUMCS, AUTONUM_CS_YARD=@AUTONUM_CS_YARD,VOLUMES=@Quantidade,dt_associacao=getdate() where autonum=@Marcante;
		END 
		IF @FlagChronos=0
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CARGA_AC=@AUTONUMCS, AUTONUM_CS_YARD_AC=@AUTONUM_CS_YARD,VOLUMES=@Quantidade,dt_associacao=getdate() where autonum=@Marcante;
		END 
	END 
	INSERT INTO SGIPA..TB_HIST_SHIFTING_CS(MARCANTE,ARMAZEM,PATIO,YARD,DT_MOV,USUARIO) VALUES (@Marcante,0,@Patio,@Yard,Getdate(),@Usuario);
	
END
GO 


CREATE Procedure pcMovimentaCS2(@Tag nvarchar(1),@Tipo nvarchar(1),@FlagChronos numeric(1),@AutonumCs numeric(10)
,@Patio numeric(2),@Yard nvarchar(10), @Quantidade numeric(8), @Motivo numeric(8)
,@Marcante numeric(12), @Usuario numeric(8),@FlagExp numeric(1)
,@Armazem numeric(8), @Fl_Frente numeric(1), @Fl_Fundo numeric(1), @Fl_LE numeric(1),@Fl_LD numeric(1)
,@AutonumCY numeric(8), @YardAtual nvarchar(30), @FlagCT numeric(1), @PercentCt numeric (6)
) 
WITH 
EXECUTE AS CALLER
AS
BEGIN
	Declare @Autonum_Cs_Yard numeric(10)
	SET @Autonum_Cs_Yard=0;
	if @FlagChronos=1 AND @FlagExp=0
	BEGIN
		INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD(AUTONUM_CS,PATIO,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD)
		VALUES (@AutonumCS,@Patio,@Armazem,@Yard,'I',@Quantidade,@Motivo,0,@FL_FRENTE,@FL_FUNDO,@FL_LE,@FL_LD);
		SELECT @Autonum_Cs_Yard = IDENT_CURRENT('TB_CARGA_SOLTA_YARD');
		IF @Tipo='Y' and @AutonumCY>0 
		BEGIN
			UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE-@QUANTIDADE WHERE AUTONUM=@AutonumCY;
		END
	END
	IF @FlagChronos=0 AND @FlagExp=0
	BEGIN
		INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD_AC(AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD)
		VALUES (@AutonumCS,@Armazem,@Yard,'I',@Quantidade,@Motivo,0,@FL_FRENTE,@FL_FUNDO,@FL_LE,@FL_LD);
		SELECT @Autonum_Cs_Yard = IDENT_CURRENT('TB_CARGA_SOLTA_YARD_AC');
		IF @Tipo='W' and @AutonumCY>0 
		BEGIN
			UPDATE SGIPA..TB_CARGA_SOLTA_YARD_AC SET QUANTIDADE=QUANTIDADE-@QUANTIDADE WHERE AUTONUM=@AutonumCY;
		END
	END 
	IF @FlagExp=1
	BEGIN
		INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD_CEXP(AUTONUM_CEXP,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD)
		VALUES (@AutonumCS,@Armazem,@Yard,'I',@Quantidade,@Motivo,0,@FL_FRENTE,@FL_FUNDO,@FL_LE,@FL_LD);
		SELECT @Autonum_Cs_Yard = IDENT_CURRENT('TB_CARGA_SOLTA_YARD_CEXP');
		IF @Tipo='T' and @AutonumCY>0 
		BEGIN
			UPDATE SGIPA..TB_CARGA_SOLTA_YARD_CEXP SET QUANTIDADE=QUANTIDADE-@QUANTIDADE WHERE AUTONUM=@AutonumCY;
		END
		IF @Yard='CAM'
		BEGIN
			UPDATE SGIPA..TB_CARGA_EXPORTACAO SET QUANTIDADE_SAIDA=ISNULL(QUANTIDADE_SAIDA,0) + @QUANTIDADE WHERE AUTONUM=@AutonumCs;
			UPDATE SGIPA..TB_CARGA_EXPORTACAO SET DT_SAIDA=GETDATE() WHERE ISNULL(QUANTIDADE_SAIDA,0) >=QUANTIDADE AND AUTONUM=@AutonumCs;
		END 
		if RIGHT(@YardAtual,3)='CAM' AND @YARD<>'CAM'
		BEGIN
			UPDATE SGIPA..TB_CARGA_EXPORTACAO SET QUANTIDADE_SAIDA=ISNULL(QUANTIDADE_SAIDA,0) - @QUANTIDADE WHERE AUTONUM=@AutonumCs;
            UPDATE SGIPA..TB_CARGA_EXPORTACAO SET DT_SAIDA=NULL WHERE ISNULL(QUANTIDADE_SAIDA,0)<QUANTIDADE AND AUTONUM=@AutonumCs;
		END
	END
	IF @Tag<>'M' AND @Marcante>'0'
	BEGIN
		IF @FlagChronos=1 AND @FlagExp=0
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CARGA=@AUTONUMCS, AUTONUM_CS_YARD=@AUTONUM_CS_YARD,VOLUMES=@Quantidade,dt_associacao=getdate() where autonum=@Marcante;
		END 
		IF @FlagChronos=0 AND @FlagExp=0
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CARGA_AC=@AUTONUMCS, AUTONUM_CS_YARD_AC=@AUTONUM_CS_YARD,VOLUMES=@Quantidade,dt_associacao=getdate() where autonum=@Marcante;
		END 
		IF @FlagExp=1
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CEXP=@AUTONUMCS, AUTONUM_CS_YARD_CEXP=@AUTONUM_CS_YARD,VOLUMES=@Quantidade,dt_associacao=getdate() where autonum=@Marcante;
		END 
	END 
	IF @TAG='M' 
	BEGIN
		IF @FlagChronos=1
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(),ARMAZEM_REC=@ARMAZEM,YARD_REC=@YARD,AUTONUM_CARGA=@AUTONUMCS, AUTONUM_CS_YARD=@AUTONUM_CS_YARD WHERE AUTONUM=@MARCANTE;
		END
	    IF @FlagChronos=0
		BEGIN
			UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(),ARMAZEM_REC=@ARMAZEM,YARD_REC=@YARD,AUTONUM_CARGA_AC=@AUTONUMCS, AUTONUM_CS_YARD_AC=@AUTONUM_CS_YARD WHERE AUTONUM=@MARCANTE;
		END
	END 
	INSERT INTO SGIPA..TB_HIST_SHIFTING_CS(MARCANTE,ARMAZEM,PATIO,YARD,DT_MOV,USUARIO) VALUES (@Marcante,@Armazem,@Patio,@Yard,Getdate(),@Usuario);
	IF @FLAGCT=1 
	BEGIN
		UPDATE SGIPA..TB_ARMAZENS_IPA SET PERC_OCUPACAO_CT=@PercentCt WHERE AUTONUM=@Armazem;
	END
END
GO 
