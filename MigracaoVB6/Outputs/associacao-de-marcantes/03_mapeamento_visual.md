# Template 03 - Mapeamento Visual

## Inventario de componentes
- Labels:
  - titulo `ASSOCIACAO DE MARCANTES`
  - `LOTE`, `CNTR`, `EXP`, `MARCANTE`, `QTDE`
  - labels auxiliares de atalhos (`[F1]-ADICIONA`, `[F2]-APAGA`, `[F4]-FILTRA`)
- Inputs:
  - `TXTLOTE`
  - `txtMarcante`
  - `txtQtdeMarcante`
- Combos:
  - `dcConteiner` (DataCombo ligado ao `AdodcCntr`)
- Grid/ListView:
  - `lv` (colunas lote/item/qtd/embalagem/conteiner/mercadoria + IDs ocultos)
- Listas:
  - `lstMarcantes` (lista temporaria de marcantes)
- Botoes:
  - `[F4] FILTRAR`, `[F1] ADICIONA`, `[F2] REMOVE`, `[F5] SALVAR`, `[ESC] SAIR`
- Secoes/Paineis:
  - frame `frmMarcantes` para area de inclusao/lista
  - area principal de filtro + grid

## Equivalencia VB6 -> MVC
- Controle VB6: `TXTLOTE`
  - Equivalente web: `<input type="text" class="form-control">`
  - Observacao: manter comportamento de disparar recarga de CNTR ao perder foco (ou no clique de filtrar).
- Controle VB6: `dcConteiner`
  - Equivalente web: `<select>` com carga dinamica (Ajax) por lote.
  - Observacao: no VB usa ADODC; no MVC substituir por endpoint JSON.
- Controle VB6: `lv` (ListView)
  - Equivalente web: tabela HTML responsiva com linha selecionavel.
  - Observacao: manter colunas tecnicas ocultas no DOM/data attributes (`autonumCs`, `autonumCntr`).
- Controle VB6: `LBLEXPORTACAO`
  - Equivalente web: badge/label condicional "EXP".
  - Observacao: estado precisa vir do backend de filtro.
- Controle VB6: `lstMarcantes`
  - Equivalente web: tabela/lista de staging com item selecionavel para remover.
  - Observacao: manter formato exibido e separacao por marcante/qtde.
- Controle VB6: botoes com atalhos
  - Equivalente web: botoes `button` + binding de `keydown`.
  - Observacao: tratar conflito de F5 (refresh do browser).

## Proposta de layout web
- Estrutura geral:
  - cabecalho da tela
  - linha de filtros (lote, cntr, botao filtrar)
  - grid principal de cargas
  - secao de associacao (marcante, qtde, botoes adicionar/remover, lista temporaria)
  - rodape de acoes (salvar/sair)
- Agrupamento de secoes:
  - manter semantica visual do legado para reduzir curva de aprendizado operacional.
- Hierarquia visual:
  - dar destaque para linha selecionada do grid e status EXP.

## Comportamentos visuais dinamicos
- Exibicao condicional:
  - badge `EXP` aparece somente quando a fonte ativa for `VW_WMS_SEM_MARCANTE_CEXP`.
- Estado de habilitado/desabilitado:
  - no VB6 `cmdAdiciona/cmdRemove` iniciam desabilitados e habilitam quando ha resultado no grid.
- Selecao e foco:
  - ao selecionar linha, recarregar lista temporaria.
  - apos adicionar, focar marcante.
  - apos salvar com sucesso, focar lote.

## UX e acessibilidade
- Melhorias sugeridas:
  - mostrar validacoes em linha sem perder mensagens equivalentes do legado.
  - confirmar selecao da linha antes de permitir adicionar/remover.
- Riscos de usabilidade:
  - atalhos globais podem conflitar com browser.
  - ausencia de setas laterais no codigo, apesar da imagem de referencia.
- Atalhos e navegacao por teclado:
  - manter F1/F2/F4/F5/ESC com `preventDefault` controlado.
  - fallback por clique para navegadores/dispositivos sem teclado funcional.

## Observacoes da imagem de referencia
- Campos e botoes observados na imagem batem com o `Marcantes.frm` (Lote, CNTR, EXP, Marcante, Qtde, F1/F2/F4/F5/ESC, lista).
- Ha setas laterais na imagem que nao aparecem explicitamente no codigo do form.
- Hipotese: setas podem vir de outro componente/versao/asset nao textual do VB6.
