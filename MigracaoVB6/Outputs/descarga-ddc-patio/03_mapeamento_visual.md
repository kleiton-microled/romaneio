# Mapeamento visual — Descarga DDC (Patio)

## Visao geral do layout (VB6)

- Janela **fixa**, sem maximizar (`BorderStyle = Fixed Single`), fundo verde-acinzentado (`&H00BA9F8D&`).
- **Faixa superior centralizada** (`Label5`): texto **"CARREGAMENTO DDC"** em destaque (contradicao com titulo da janela "Descarga Armazem" — tratar como divida de UX na migracao).
- Area superior: selecao **Placa** (`cbPlaca`) e **Lote** (`cbLote`).
- Acoes superiores: **F4 - Carregar**, **F2 - Limpar**.
- **Frame** inferior (`Frame1`), inicialmente desabilitado: contem formulario de item + **grid**.

## Inventario de componentes

| Controle VB | Texto / papel | Equivalente web sugerido |
|-------------|----------------|---------------------------|
| `Label5` | Titulo de contexto | `<h1>` ou barra de titulo da pagina + design system |
| `Label7` / `Label6` | "Placa", "Lote" | `<label>` associado a selects |
| `cbPlaca`, `cbLote` | Combos ligados a ADODC | `<select>` com busca (Select2/Choices) ou autocomplete se lista grande |
| `Command1` | F4 - Carregar | Botao primario + atalho (ver nota de atalhos) |
| `Command2` | F2 - Limpar | Botao secundario |
| `Frame1` | Agrupamento de itens | `<fieldset>` ou card (`div` com heading) |
| `Label2`, `Label3`, `Label9` | Quantidade, Embalagem, pergunta finalizacao | labels |
| `TxtQuantidade` | Entrada numerica curta | `<input type="number" min="1">` ou texto com validacao |
| `DcEmbalagem` | Combo embalagem | `<select>` carregado do servidor |
| `Option1` / `Option2` | SIM / NAO | `<input type="radio">` em grupo nomeado |
| `Grid1` | Lista de itens | `<table>` responsiva ou grid JS (DataTables) com linha selecionavel |
| `Command5`, `Command3`, `Command4` | F3 Novo, F5 Salvar, F6 Excluir | botoes com `accesskey` ou shortcuts JS documentados |

## Estados visuais dinamicos

| Estado | Evidencia |
|--------|-----------|
| Area de itens desabilitada | `Frame1.Enabled = False` ate carregar |
| Atalhos F3/F5/F6 bloqueados | `KeyDown` verifica `Frame1.Enabled` |
| Linha do grid selecionada | `Grid1_Click` copia dados para formulario de edicao |

## Atalhos de teclado no contexto web

- **Hipotese de UX**: mapear `F2/F3/F4/F5/F6` apenas quando foco estiver em container da tela e sem conflito com o browser (muitos navegadores reservam `F3`, etc.). Alternativa segura: **botoes visiveis** + atalhos opcionais com `preventDefault` controlado.
- `ESC` para fechar: na web, preferir **botao Sair** explicito; `ESC` pode fechar modal se a tela for modal.

## Agrupamentos

- **Filtro / cabecalho**: placa + lote + carregar/limpar.
- **Detalhe / operacao**: quantidade + embalagem + finalizacao + acoes + grid.

## Equivalencia Razor / HTML / CSS / JS

- View Razor com layout de **duas secoes** (cabecalho sticky opcional para operacao em campo).
- Scripts em arquivo dedicado (`Content/js/...js`) seguindo padrao de telas recentes do Romaneio (ex.: endpoints JSON no controller).
- Validacao client-side **espelhando** mensagens do legado + validacao server-side obrigatoria.

## Pontos de UX a adaptar para web

1. **Combos ADO ligados a controles ocultos**: substituir por endpoints JSON (`GET` lotes/placas, `GET` embalagens).
2. **DataGrid VB**: selecao de linha e edicao inline — na web, preferir **linha selecionada + formulario** (como o proprio legado faz ao clicar no grid).
3. **Textos inconsistentes** ("Descarga Armazem" vs "Descarga DDC" vs "CARREGAMENTO DDC"): definir com negocio um **titulo unico** na migracao sem mudar regra de dados.
4. **Campos ocultos**: na web, nao expor IDs sensiveis na UI; manter em ViewModel/estado JS ou sessao conforme padrao Romaneio.
5. **Feedback de erro de gate**: hoje e `MsgBox`; na web usar **toast/alert** acessivel (ARIA live region).

## Acessibilidade

- Opcoes SIM/NAO devem ter **legend** claro ("Entrega finalizada?").
- Atalhos: documentar na tela (`<kbd>`) e nao depender apenas de teclas de funcao.
