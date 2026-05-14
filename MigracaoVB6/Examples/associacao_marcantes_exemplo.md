# Exemplo rapido - Aplicacao dos agentes

Tela de referencia: **Associacao de Marcantes** (`ProjetoReferencia/Coletor/Marcantes.frm`)

## A01 Leitura VB6 (resumo)
- Form principal localizado: `Marcantes.frm`
- Arquivos correlatos: `Marcantes.frx`, `Principal.frm`, `Microled.bas`, `mdlColetor.bas`, `frmMensagem2.frm`
- Eventos chave: `Form_Load`, `Form_KeyDown`, `cmdFiltrar_Click`, `cmdAdiciona_Click`, `cmdRemove_Click`, `cmdSalvar_Click`, `lv_Click`
- Dependencias SQL: `VW_WMS_SEM_MARCANTE`, `VW_WMS_SEM_MARCANTE_CEXP`, `TB_TEMP_MARCANTE`, `TB_TEMP_MARCANTE_CEXP`, `TB_MARCANTES`

## A02 Mapeamento funcional (resumo)
- Objetivo: associar marcantes a cargas pendentes de marcacao
- Acoes: filtrar por lote/cntr, incluir marcante, remover, salvar associacao, sair
- Fluxo principal: selecionar carga -> montar lista temporaria -> validar quantidades -> persistir

## A03 Mapeamento visual (resumo)
- Controles principais: lote, cntr, grid de cargas, secao EXP, marcante, qtde, lista `lstMarcantes`, botoes de acao
- Equivalencia: ListView VB6 -> tabela web; lista VB6 -> lista staging em grid/lista web; atalhos -> keydown JS

## A04 Regras de negocio (resumo)
- Marcante obrigatorio e valido
- Quantidade obrigatoria
- Bloqueio de marcante duplicado na lista temporaria
- Regra de consistencia de quantidade antes de salvar
- Regra especifica para modo exportacao (EXP)

## A05 Desenho MVC (resumo)
- Proposta: `AssociacaoMarcantesController` + `IAssociacaoMarcantesRepositorio` + `AssociacaoMarcantesViewModel`
- Endpoints: filtrar, carregar temp, incluir temp, remover temp, salvar
- Scripts: controle de selecao de linha, atalhos F1/F2/F4/F5/ESC, feedback de validacao

## A06 Plano de implementacao (resumo)
1. Criar contratos e repositorio
2. Criar controller e rotas de acao
3. Construir view com grid/lista staging
4. Implementar validacoes e persistencia
5. Testar fluxos principal, erro e borda
