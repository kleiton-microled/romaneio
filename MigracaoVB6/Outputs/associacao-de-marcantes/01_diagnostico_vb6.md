# Template 01 - Diagnostico VB6

## Identificacao da tela
- Nome: Associacao de Marcantes
- Modulo: Coletor (IPA), menu principal do coletor
- Form principal: `Marcantes`
- Caminho: `ProjetoReferencia/Coletor/Marcantes.frm`

## Arquivos localizados
- `.frm`:
  - `ProjetoReferencia/Coletor/Marcantes.frm` (form alvo)
  - `ProjetoReferencia/Coletor/Principal.frm` (fluxo de entrada `Case 8 -> Marcantes.Show 1`)
  - `ProjetoReferencia/Coletor/frmMensagem2.frm` (mensagens modal)
- `.frx`:
  - `ProjetoReferencia/Coletor/Marcantes.frx` (binding do `dcConteiner`)
- `.bas`:
  - `ProjetoReferencia/Coletor/Microled.bas` (globais `db`, `ConnectionString`, `Nnull`, `SoNumeros09`)
  - `ProjetoReferencia/Coletor/mdlColetor.bas` (`Col_Mensagem`, `Col_Imagem`, `Valida_Acesso_Botao`)
- `.cls`:
  - Nenhum `.cls` especifico da tela localizado no projeto `Coletor`
- Outros:
  - `ProjetoReferencia/Coletor/Coletor.vbp` (composicao do projeto e dependencias externas `..\Sgipa\...`)
  - `ProjetoReferencia/Coletor/Marcante_Qtde.sql` (contexto de estrutura/consultas de marcante)

## Eventos mapeados
- Form_Load:
  - Configura colunas do `ListView lv` (inclui colunas ocultas de IDs: `AUTONUMCS` e `AUTONUMCNTR`)
  - Nao carrega dados automaticamente
- Form_KeyDown/KeyPress:
  - `F4` filtra (`Carrega_Lv`)
  - `F1` adiciona (`cmdAdiciona_Click`)
  - `F2` remove (`cmdRemove_Click`)
  - `F5` salva (`cmdSalvar_Click`)
  - `ESC` fecha tela (`Unload Me`)
  - `txtMarcante_KeyPress` restringe a digitos via `SoNumeros09`
- Clicks de botoes:
  - `cmdFiltrar_Click` -> `Carrega_Lv`
  - `cmdAdiciona_Click` -> valida + grava em temp
  - `cmdRemove_Click` -> remove da temp
  - `cmdSalvar_Click` -> valida consistencia + atualiza `TB_MARCANTES`
  - `cmdEscape_Click` -> sair
- Eventos de lista/grid:
  - `lv_Click` limpa lista e recarrega temp por linha selecionada
  - `lv_KeyUp` apenas limpa lista (sem recarregar temp)
  - `dcConteiner_KeyDown` com `Delete` limpa texto do combo
- Eventos de foco:
  - `TXTLOTE_LostFocus` recarrega combo de conteiner via `AdodcCntr`
  - `txtMarcante_GotFocus` limpa campo e pinta fundo
  - `txtMarcante_LostFocus` restaura cor de fundo

## SQLs e persistencia
- Consultas:
  - `SELECT ... FROM SGIPA..VW_WMS_SEM_MARCANTE` (principal)
  - fallback `SELECT ... FROM SGIPA..VW_WMS_SEM_MARCANTE_CEXP` (modo exportacao)
  - `SELECT AUTONUM,... FROM SGIPA..TB_MARCANTES WHERE AUTONUM = <marcante>`
  - `SELECT MARCANTE,QTDE FROM SGIPA..TB_TEMP_MARCANTE WHERE AUTONUM_CS = <id>`
  - `SELECT MARCANTE,QTDE FROM SGIPA..TB_TEMP_MARCANTE_CEXP WHERE AUTONUM_CEXP = <id>`
  - `SELECT QUANTIDADE_REAL FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM = <id>`
  - `SELECT QUANTIDADE - ISNULL(QUANTIDADE_SAIDA,0) FROM SGIPA..TB_CARGA_EXPORTACAO WHERE AUTONUM = <id>`
  - `SELECT SUM(VOLUMES) FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CARGA = <id>`
  - `SELECT SUM(VOLUMES) FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CEXP = <id>`
  - `SELECT DISTINCT autonumCNTR, ID_CONTEINER AS DISPLAY FROM SGIPA..VW_WMS_SEM_MARCANTE ...` (combo CNTR)
- Inserts:
  - `INSERT INTO SGIPA..TB_TEMP_MARCANTE(AUTONUM_CS,MARCANTE,QTDE)`
  - `INSERT INTO SGIPA..TB_TEMP_MARCANTE_CEXP(AUTONUM_CEXP,MARCANTE,QTDE)`
- Updates:
  - `UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(), AUTONUM_CARGA=..., VOLUMES=..., BL=..., ITEM=..., CNTR=... WHERE AUTONUM=...`
  - `UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(), AUTONUM_CEXP=..., VOLUMES=... WHERE AUTONUM=...`
- Deletes:
  - `DELETE FROM SGIPA..TB_TEMP_MARCANTE WHERE AUTONUM_CS=... AND MARCANTE=...`
  - `DELETE FROM SGIPA..TB_TEMP_MARCANTE_CEXP WHERE AUTONUM_CEXP=... AND MARCANTE=...`
- Views/Tabelas/Functions:
  - Views: `VW_WMS_SEM_MARCANTE`, `VW_WMS_SEM_MARCANTE_CEXP`
  - Tabelas: `TB_MARCANTES`, `TB_TEMP_MARCANTE`, `TB_TEMP_MARCANTE_CEXP`, `TB_CARGA_SOLTA`, `TB_CARGA_EXPORTACAO`
  - Functions/Procedures: nenhuma chamada explicita na tela

## Dependencias tecnicas
- Variaveis globais:
  - `db` (conexao ADODB)
  - `ConnectionString`
  - `Col_Mensagem`, `Col_Imagem`
- Funcoes utilitarias:
  - `Nnull`
  - `SoNumeros09`
- Formularios auxiliares:
  - `frmMensagem2` (mensagens/confirmacoes visuais)
- Controle de permissao:
  - Entrada da tela depende de permissao no menu (`Valida_Acesso_Botao(Col_Login, "COMMAND1(8)")` em `Principal.frm`)

## Atalhos e automatismos
- Atalhos:
  - F1 adiciona, F2 remove, F4 filtra, F5 salva, ESC sai
- Foco automatico:
  - apos adicionar: foco volta para `txtMarcante`
  - apos salvar sucesso: foco volta para `TXTLOTE`
- Comportamentos implicitos:
  - `LBLEXPORTACAO.Visible=True` sinaliza modo CEXP
  - `On Error Resume Next` em `Carrega_Marcantes_Temp` mascara falhas silenciosamente
  - `Form_Unload` chama `Unload Me` novamente

## Inventario dos objetos de banco usados pela tela
- `SGIPA..VW_WMS_SEM_MARCANTE` - carga sem marcante (fluxo normal)
- `SGIPA..VW_WMS_SEM_MARCANTE_CEXP` - carga sem marcante (fluxo exportacao)
- `SGIPA..TB_MARCANTES` - valida marcante e grava associacao final
- `SGIPA..TB_TEMP_MARCANTE` - staging temporario (normal)
- `SGIPA..TB_TEMP_MARCANTE_CEXP` - staging temporario (exportacao)
- `SGIPA..TB_CARGA_SOLTA` - consulta `QUANTIDADE_REAL`
- `SGIPA..TB_CARGA_EXPORTACAO` - consulta quantidade disponivel

## Pontos criticos
- Riscos:
  - SQL dinamico com concatenacao de string (sem parametros)
  - ausencia de validacao robusta para selecao no `lv`/`lstMarcantes` antes de operar
  - tratamento de erro silencioso em carga de temp (`On Error Resume Next`)
- Inconsistencias:
  - filtro de lote usa `Val(...)` no fluxo normal e string no fluxo CEXP (`LOTE='...'`)
  - `Form_Unload` recursivo (`Unload Me` dentro do proprio evento)
  - `lv_KeyUp` limpa lista sem recarregar temp, comportamento diferente do clique
- Hipoteses:
  - limpeza de `TB_TEMP_*` apos salvar pode ocorrer fora deste form (hipotese; nao evidenciado no codigo da tela)
  - setas laterais vistas na imagem podem ser de outro build/recurso visual nao explicitado no `.frm` (hipotese)
