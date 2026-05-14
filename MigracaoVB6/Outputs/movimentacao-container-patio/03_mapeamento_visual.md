# 03 — Mapeamento visual — Movimentação Container: Pátio

**Fonte principal:** `FrmMovimentacaoS.frm` (layout compacto do **Coletor Pátio**).  
**Imagem de referência:** não foi fornecida neste estudo; o inventário abaixo deriva do `.frm`.

---

## Layout geral

- **Form:** sem borda (`BorderStyle = 0`), não aparece na barra de tarefas; ocupa área embutida no shell do coletor.
- **Navegação:** `SSTab` com **4 abas** na variante S — **Moviment.** | **Histórico** | **Dados CNTR** | **Avarias**.
- **Fonte base:** Tahoma; abas com fonte reduzida (6pt) na variante S para caber em tela pequena.

---

## Aba **Moviment.** (núcleo operacional)

| Componente VB6 | Descrição visual | Equivalência web sugerida |
|----------------|------------------|---------------------------|
| `txtPatio` | Descrição resumida do pátio (readonly, fundo cinza) | `<input readonly>` ou texto estático no cabeçalho |
| `CNTR` | Entrada do **final** da sigla (grande, centrado) | Campo mascarado + busca ao digitar (debounce) ou botão “Buscar” |
| `MskCntr` | ID completo mascarado | Input com máscara ISO 6346 / padrão legado |
| Labels dinâmicos | `LBLDTA`, `lblDestinadoDesova`, `lblReeferDesligado`, `Label1(9)` alertas | Badges/chips coloridos (Bootstrap `badge`) |
| `TXTYardAtual` | Yard atual (readonly) | Texto destacado + ícone “pin” |
| `mskDestino` | Destino editável | Input texto uppercase; autocomplete com yards válidos |
| `DataCombo2` | Motivo (lista) | `<select>` pesquisável (Select2 já usado no Romaneio) |
| `Check1` | “LACRES CONFERIDOS” | Checkbox explícito com label forte |
| Campos de lacre / tara / GMCI | Grupo denso | `fieldset` “Lacres / Pesos” com grid responsivo |
| `Grid1` | Próximos movimentos (MSHFlexGrid) | Tabela HTML compacta (`DataTables` ou tabela simples) |
| `Grid2` | Pilha / vizinhos (MSHFlexGrid) | Tabela com ordenação por coluna PILHA |
| `Option1(0|1)` | “ATUAL” / “DESTINO” | Radio group ou toggle segmentado |
| `Command1` | “Salvar” | Botão primário “Confirmar movimentação” |
| `Command2` | “Limpar” | Botão secundário |
| `Command4` | “X” | Fechar / voltar (ícone + aria-label) |
| `Frame2` + `DataCombo4` + `Command3` | Overlay interno de **câmera** | Modal Bootstrap (`modal`) com select de câmera + confirmar |
| `Frame5` | Quadro contextual (transporte interno) | Alert `alert-info` colapsável |
| `Text1(0)` | Bloco textual multi-linha | `<pre>` ou card com scroll para observações do sistema |

---

## Aba **Histórico**

| Componente | Descrição | Web |
|------------|-----------|-----|
| `DataGrid1` | Histórico de movimentações (colunas Data, De, Para, Usuario, Equipamento, Motivo) | Tabela paginada + ordenação por data |
| `GridHist_Avaria` | Linhas de avarias com data | Sub-tabela ou lista secundária |

**Carga:** ao selecionar aba, código atualiza `Adodc1` com SQL dinâmico conforme tipo de sistema (`I`, `A`, `R`).

---

## Aba **Dados CNTR**

| Componente | Descrição | Web |
|------------|-----------|-----|
| `Frame1` “DADOS ENTRADA” | Lacres de entrada, dados físicos | Cards em duas colunas (desktop), uma coluna (mobile) |
| `DataGrid2` | BLs vinculados | Tabela somente leitura |
| Campos `txtTam`, `txtTipo`, `TxtNavio`, etc. | Atributos da unidade | Labels + valores |
| `txtSIGVIG` | Situação SIGVIG | Badge + tooltip explicativo |

---

## Aba **Avarias**

| Componente | Descrição | Web |
|------------|-----------|-----|
| `ListLocal` / `ListTipo` | Listas de seleção múltipla (com validação de único local) | Dois `<select multiple>` ou picker com chips |
| `cmdIncluirAvaria`, `cmdExcluirAvaria`, `cmdFinalizarAvaria` | Ações | Botões com confirmação (`frmMensagem2` → modal) |
| `GridAvarias` | Avarias abertas (`VW_EGMCI_AVARIAS_TTL`) | Tabela; realçar linhas com `autonum=0` (cor de fundo no VB — “apontada pelo operador”) |

**Permissão:** se usuário não tem função 998, aba pode ficar desabilitada (`TabEnabled(3)=False`).

---

## Comportamentos visuais acoplados à regra

| Comportamento | Observação UX web |
|---------------|-------------------|
| Enter → Tab | Opcional no web; pode confundir em textarea — limitar a campos curtos |
| Busca ao `Change` do `CNTR` | Substituir por busca explícita ou debounce para evitarRequests excessivos |
| Frame de câmera sobreposto | Sempre modal com foco preso e ESC para cancelar |
| Mensagens modais encadeadas | Consolidar em toasts + modal para confirmações críticas (SAÍDA, IMO) |

---

## Atalhos de teclado (legado)

- **Enter:** avança foco (emula Tab).
- Demais atalhos **não mapeados** de forma explícita neste form além de `KeyPreview` — **hipótese**: atalhos globais podem existir no shell MDI/outro hook; não evidenciado neste `.frm`.

---

## Pontos de UX para adaptação web

1. **Densidade:** a versão S empilha muita informação em poucos pixels — no web usar **wizard** ou **accordions** (Movimentação → Detalhes → Avarias).
2. **Dupla entrada CNTR:** separar claramente “final” vs “ID completo” com feedback de ambiguidade (igual mensagens do legado).
3. **Grids simultâneos:** em mobile, usar abas internas “Próximos movimentos” / “Pilha”.
4. **Confirmação de SAÍDA e IMO:** manter **dupla confirmação** visual (modal destacado).
5. **Acessibilidade:** garantir leitura de ordem lógica e `aria-live` para mensagens de bloqueio Reefer/scanner.

---

## Hipóteses

1. A variante `FrmMovimentacao.frm` (não-S) pode ter pequenas diferenças de posicionamento — conferir antes de reproduzir pixel-perfect.
2. Ícones das abas vêm do `.frx`; não há garantia de equivalência sem asset exportado.
