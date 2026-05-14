# Mapeamento visual - Manutencao Marcantes (AdminMarcante)

## Fonte visual

- **Layout**: propriedades do form em `AdminMarcante.frm` (VB6).
- **Imagem de referencia**: **nao** foi localizada imagem anexa neste estudo; somente o `.frm`.

## Visao geral do layout

- Form **sem borda** (`BorderStyle = 0`), **sem caixa de controle**, `KeyPreview = True`.
- Fundo do form: cor `&H00BA9F8D&` (verde acinzentado).
- Area compacta (~3540 x 4455 twips no designer), orientada a coletor / resolucao baixa.

## Agrupamentos logicos (de cima para baixo)

1. **Cabecalho**: titulo `MANUTENCAO- MARCANTES` (label superior em faixa clara).
2. **Filtro por lote**: rotulo "LOTE" + `TXTLOTE` + botao `[F4]-FILTRA`.
3. **Resumo da carga / divergencia**: "EMBALAGEM" + `txtEmbalagem`; "QTDE" + `txtQtde` (quantidade de referencia).
4. **Lista de marcantes**: titulo "MARCANTES" + `lv` (ListView com gridlines).
5. **Totalizador**: rotulos "TTL" / `TXTTOT` ao lado da soma.
6. **Edicao**: "MARCANTE" + `TXTMARCANTE`; "QTDE" + `TXTQTDEN` (nova quantidade).
7. **Rodape**: `[F5]-SALVA` e `[ESC]-VOLTAR`.

## Inventario de componentes VB6 -> equivalencia web (Razor / HTML / JS)

| VB6 | Equivalencia sugerida | Observacoes |
|-----|----------------------|-------------|
| `Label1`, `Label3` | `<label>` ou texto em `<div class="form-label">` | Manter hierarquia semantica. |
| `TXTLOTE` | `<input type="text" inputmode="numeric" maxlength="7">` | Alinhar `maxlength` ao legado; validar inteiro no servidor. |
| `txtEmbalagem`, `txtQtde`, `TXTTOT`, `TXTMARCANTE` (somente leitura onde `Locked=True`) | `<input readonly>` ou texto estático | `TXTMARCANTE` no legado nao esta `Locked`; apos selecao na grid, pode ser readonly na web para evitar edicao manual indevida — **Hipotese de UX** (nao altera paridade se espelhar VB6). |
| `TXTQTDEN` | `<input type="number">` ou texto numerico com validacao | Evitar valores negativos na web — **melhoria segura** se negocio permitir (nao evidenciado no VB6). |
| `lv` (ListView) | `<table>` responsiva ou componente de grid (DataTables, grid interno do projeto) | Colunas visiveis conforme `Form_Load`; coluna `AUTONUMCS` oculta -> coluna `d-none` ou dado em `data-*`. |
| `cmdFiltra`, `cmdSalvar`, `cmdEscape` | `<button type="button">` | Manter rotulos com atalhos ou tooltips. |
| Cor dinamica `TXTTOT` (verde/vermelho/preto) | classes CSS (`text-success`, `text-danger`) | Estado: igual referencia (verde), diferente (vermelho), limpo (preto). |

## Atalhos exibidos na UI

- `[F4]-FILTRA`, `[F5]-SALVA`, `[ESC]-VOLTAR`.

**Adaptacao web**: registrar listeners em `keydown` no container da pagina; documentar conflito com atalhos do navegador (F5 recarregar pagina) — ja observado em outras migracoes do Romaneio.

## Estados visuais dinamicos

| Estado | Evidencia |
|--------|-----------|
| Total igual referencia | `TXTTOT.ForeColor = vbGreen` quando `TXTTOT = TXTQTDE`. |
| Total diferente da referencia | `TXTTOT.ForeColor = vbRed`. |
| Estado "limpo" | `TXTTOT.ForeColor = vbBlack` em `Limpa`. |
| Campo `TXTAUTONUMCS` | `Visible = False` — na web, campo oculto ou estado em sessao/viewmodel. |

## UX e acessibilidade

- **Selecao na lista**: na web, preferir clique em linha + estado visual de linha selecionada (equivalente a `lv_Click`).
- **Teclado**: garantir ordem de tabulacao logica (lote -> filtrar -> lista -> quantidade -> salvar).
- **Mensagens modais**: substituir `frmMensagem2` por toasts/modais Bootstrap (padrao comum no Romaneio) mantendo textos.
- **Feedback de gravacao**: o legado **nao** exibe mensagem de sucesso apos `db.Execute` — **Lacuna de UX** para definir na migracao (toast de sucesso/erro).

## Checklist interno (Agente 03)

- [x] Controles mapeados.
- [x] Equivalente web definido por controle.
- [x] Estados visuais dinamicos mapeados.
- [x] Atalhos e proposta web.
- [x] Riscos de usabilidade documentados.
