# Mapeamento Visual - Carregamento Carga Solta (Armazem)

## Inventario de componentes

### Cabecalho
- `Label2(32)` - barra de titulo "CARREGAMENTO CARGA SOLTA" (BorderStyle FixedSingle, fonte Tahoma 6 bold).

### Selecao de veiculo
- `Label1(0)` - "VEIC.".
- `DataCombo1` - combo (style 2 = DropDownList) com placas `PLACA_C + ' ' + PLACA_CARRETA + ' - ' + MODELO` vindas de `VW_CAM_CARREGAMENTO`.
- `Command1` - botao `...` (Refresh) - recarrega o combo.
- `Adodc1` - data control oculto que alimenta o `DataCombo1`.

### Lista de Ordens
- `Label3(0)` - "ORDENS" (cabecalho da secao).
- `lvOC` - ListView (View=Report) - colunas:
  - `none` (oculta) - chave da linha.
  - `#ORDEM` - `NUM_OC`.
  - `LOTE` - `LOTE` (BL.AUTONUM).
  - `ITEM` (oculta) - `ITEM`.
  - `QTDE` - `QTDE_CARREGADA/QUANTIDADE`.
  - `EMBALAGEM` - `EMBALAGEM`.
  - `OC` (oculta) - `ORDEM_CARREG`.
  - `CS` (oculta) - `AUTONUMCS`.

### Lista de Carregamentos
- `Label3(1)` - "CARREGAMENTOS" (cabecalho da secao).
- `lvCm` - ListView (View=Report) - colunas:
  - `NONE` (oculta).
  - `MARCANTE` - `M.AUTONUM` (formatado com 9 digitos).
  - `QTDE` - `M.VOLUMES`.

### Bipagem do marcante
- `Label1(1)` - "MARCANTE".
- `txtMarcante` - input numerico, `MaxLength=9`, `IMEMode=disable`, BackColor amarelo claro.
- `Label1(2)` - "LOTE".
- `TXTLOTE` - input readonly (Locked=True, BackColor cinza), recebe `BL` da carga.
- `Label1(3)` - "QTDE".
- `TXTQTDE` - input editavel para quantidade carregada (default = volumes do marcante; usuario pode reduzir para parcial). `Tag` armazena valor original.
- `Label1(6)` - "LOCAL".
- `TXTLOCAL` - input readonly mostrando `Y.YARD` (posicao em yard) - quando vazio, marcante esta no armazem; quando `'CAM'`, ja foi carregado.
- `txtArmazem` - input readonly mostrando `DESCR_ARMAZEM`. `Tag` guarda `AUTONUM_ARMAZEM`.
- `txtAutonumCS` - hidden (Visible=False, MaxLength=9). Guarda `AUTONUMCS` da carga.
- `txtAutonum_Cs_Yard` - hidden. Guarda `AUTONUM_CS_YARD` (id da posicao yard atual do marcante).

### Mensagens / critica
- `txtCritica` - textbox readonly, BackColor preto, ForeColor verde/vermelho - mensagem operacional.

### Botoes principais
- `cmdF1` - "[F1] LOCAIS" + `Label3(3)` "[F1]-LOCAIS".
- `cmdF4` - "[F4] CARREGA" + `Label1(4)` "[F4]-CARREGA ITEM".
- `Command2` - "[F3] LIMPAR" (azul/vermelho).
- `cmdF2` - "[F2] DOC".
- `cmdF5` - "[F5] CANCELA" + `Label3(2)` "[F5]-CANCELA".

### Painel sobreposto LOCAIS (`frmLocais`, Visible=False inicialmente)
- `Label4(0)` - "Ordem de Carregamento :".
- `txt2Ordem` - input readonly (preenchido com `ORDEM_CARREG` da OC selecionada).
- `Label4(1)` - "Lote:".
- `txt2Lote` - input readonly (preenchido com `LOTE` da OC selecionada).
- `lvLocal` - ListView (Report) - colunas:
  - `NONE` (oculta).
  - `MARCANTE` - formato 9 digitos.
  - `QTDE`.
  - `LOCAL` - `DESCR_ARMAZEM + " " + POSICAO`.

## Equivalencia VB6 -> MVC

| Controle VB6 | Equivalente web (Razor/Bootstrap/JS) | Observacao |
|---|---|---|
| `DataCombo1` (DataCombo) | `<select id="ddlVeiculo" class="form-control">` populado por endpoint `Veiculos`/preload server-side | Mesmo conteudo, sem ADO control. |
| `Command1` ("...") | `<button id="btnRefreshVeiculos" class="btn btn-light">Atualizar</button>` ou icone refresh ao lado do select | Re-aciona endpoint do combo. |
| `lvOC` (ListView Report) | `<table id="tblOrdens" class="table table-sm table-hover">` com selecao por linha (clique/teclado) | Atributos `data-autonumcs`, `data-ordem`, `data-lote`, `data-quantidade-original` etc. |
| `lvCm` (ListView Report) | `<table id="tblCarregamentos" class="table table-sm">` somente leitura | Atualizado apos cada `F4`/`F5`. |
| `txtMarcante` | `<input id="txtMarcante" type="text" inputmode="numeric" maxlength="9" pattern="\d*" autocomplete="off" autofocus>` | JS captura `keydown` e `input` para implementar deteccao de bipagem (delta de tempo). |
| `TXTLOTE` / `txtArmazem` / `TXTLOCAL` | `<input class="form-control" readonly>` | Apenas leitura. |
| `TXTQTDE` | `<input id="txtQtde" type="number" min="0" class="form-control">` | Atributo `data-qtde-original` substitui `Tag` do VB6. |
| `txtAutonumCS` / `txtAutonum_Cs_Yard` | `<input type="hidden">` (ou state em JS) | Mantidos para evitar idas extras ao banco. |
| `txtCritica` | `<div id="critica" class="alert" role="status" aria-live="polite">` | Cor verde para sucesso, vermelha para erro. Substitui textbox preto/verde do VB6. |
| `cmdF1` LOCAIS | `<button data-action="locais">[F1] Locais</button>` | Abre modal/offcanvas substituto do `frmLocais`. |
| `cmdF2` DOC | `<button data-action="doc">[F2] Doc</button>` | Abre nova aba/janela web (substituicao do exe externo - ver Decisoes). |
| `cmdF4` CARREGA | `<button class="btn btn-success" data-action="carrega">[F4] Carrega Item</button>` | POST `Carregar`. |
| `cmdF5` CANCELA | `<button class="btn btn-warning" data-action="cancela">[F5] Cancela</button>` | POST `Cancelar`. |
| `Command2` LIMPAR | `<button class="btn btn-danger" data-action="limpa">[F3] Limpar</button>` | Reseta campos do marcante. |
| `frmLocais` (frame interno) | `<div class="modal fade" id="modalLocais">` (Bootstrap modal) ou `offcanvas` | ESC fecha e devolve foco. |
| `lvLocal` | `<table id="tblLocais" class="table table-sm">` dentro do modal | Mesmas colunas. |

## Proposta de layout web
- Estrutura geral (uma unica view `Index.cshtml` no padrao das outras telas do Romaneio):
  - Cabecalho com titulo "Carregamento Carga Solta - Armazem" e icone (`glyphicon-truck` ou similar), seguindo o padrao das demais telas (`MovimentacaoCS`, `AssociacaoMarcantes`).
  - Linha 1 (Veiculo): combo `Veiculo` ocupando 8 colunas, botao `Atualizar` (1 col), area de status compacto.
  - Linha 2 (Ordens): tabela `Ordens` (12 cols) com altura limitada e scroll, mostrando `#Ordem | Lote | Qtde Carregada/Qtde | Embalagem`.
  - Linha 3 (Carregamentos): tabela `Carregamentos` (12 cols), menor que a anterior.
  - Linha 4 (Marcante): grid responsivo com `Marcante (3 cols)`, `Qtde (2)`, `Lote (2)`, `Local (2)`, `Armazem (3)`.
  - Linha 5 (Critica): banner full-width.
  - Linha 6 (Botoes): grupo de botoes alinhados (Locais | Doc | Limpar | Carrega | Cancela | Sair).
  - Modal Locais: tabela centralizada com `Marcante | Qtde | Local`. Botao fechar e ESC.
- Agrupamento de secoes:
  - "Selecao de Veiculo".
  - "Ordens de Carregamento".
  - "Itens ja carregados (CAM)".
  - "Bipagem".
  - "Acoes".
- Hierarquia visual:
  - Indicador de progresso "QTDE_CARREGADA / QUANTIDADE" como badge com cor (verde quando igual, ambar quando parcial, cinza quando zero).
  - Linha selecionada de `lvOC` com destaque (azul claro) - reaproveita estilo das outras tabelas.
  - Modal de locais com header coerente com `_Layout.cshtml`.

## Comportamentos visuais dinamicos
- Exibicao condicional:
  - Botoes `F1`, `F2`, `F4`, `F5` so ficam habilitados apos o usuario selecionar uma OC e/ou bipar marcante (regras especificas, equivalentes a logica do VB6).
  - Modal de Locais inicia oculto; abre via `F1` ou clique no botao; fecha via `ESC` ou click fora.
- Estado de habilitado/desabilitado:
  - `txtMarcante` desabilita apos sucesso quando `Wflag_carregamento_marcante_bip = 1` (replica VB6) - controlado por flag retornada pelo endpoint que monta a tela.
  - `TXTQTDE` permanece editavel quando ha `data-qtde-original > 0`, senao readonly (zero).
- Selecao e foco:
  - Apos `F4`/`F5` sucesso, foco volta para `txtMarcante` (consistente com VB6).
  - Apos abrir modal Locais, foco entra na primeira linha da tabela; `ESC` retorna foco ao input marcante.
  - Apos refresh do combo, mantem item selecionado se ainda existir.

## UX e acessibilidade
- Melhorias sugeridas:
  - Mostrar linha "Carregado/Total" como progress bar inline na lista de ordens (ja existe no VB6 como texto `n/m`).
  - Adicionar campo de busca rapida na lista de ordens (filtra por lote/cntr).
  - Substituir `txtCritica` (textbox preto) por banner Bootstrap com `role=status`/`aria-live=polite` para leitores de tela.
  - Mostrar tooltip sobre `MARCANTE` lembrando que so aceita 9 digitos.
- Riscos de usabilidade:
  - Operadores estao acostumados aos atalhos `F1/F2/F3/F4/F5`; manter mapeamento exato em JS (`event.key === 'F4'`, `e.preventDefault()`).
  - Em alguns navegadores, `F1` abre ajuda; usar `preventDefault` no `keydown` do form.
  - Anti-digitacao manual de marcante: precisa de delta de tempo confiavel; testar com leitores reais (USB HID enviam `Enter` ao final).
  - Diferenca entre `LOCAL='CAM'` (carregado) e vazio (no armazem) deve ficar visualmente clara - sugerimos badge.
- Atalhos e navegacao por teclado:
  - `F1` Locais, `F2` Doc, `F3` Limpar, `F4` Carrega, `F5` Cancela, `ESC` Sair / Fechar Modal.
  - `Enter` em `txtMarcante` aciona `Carrega_Dados` (replica `KeyPress=13` do VB6) quando `flag_bip = 0`; quando `flag_bip <> 0`, somente bipagem dispara.
  - `Tab` segue ordem natural: Veiculo -> Tabela Ordens -> Marcante -> Qtde -> Botoes.
  - Atalhos somente acionam se nenhum input estiver com IME aberto / nenhum modal alheio estiver na frente.
