# Mapeamento visual — Container REEFER (Patio)

## Referencia de imagem

- **Nao** foi localizada imagem estatica (PNG/JPG) no repositorio para esta tela. Evidencia visual deriva do layout `.frm` + `.frx` (icones das abas).

## Estrutura geral

- **Form** sem borda (`BorderStyle = 0`), sem barra de titulo nativa padrao de MDI filho modal.
- **Abas horizontais** (`SSTab1`) com tres secoes:
  0. **Monitoramento** (operacao principal)
  1. **Previsao** (texto no arquivo pode aparecer corrompido como "Previsуo")
  2. **Un. Desligadas**

## Aba "Monitoramento" (ReeferS — layout compacto)

### Cabecalho

- Rotulo **"CONTEINER"** (`Label1(0)`) em destaque (cor vermelha no tema VB `&H00FF0000`).
- **Alerta** `lblReeferDesligado` ("REEFER DESLIGADO") e `lblPlugOFF` (texto dinamico "PLUG OFF: ...") — inicialmente ocultos.
- **Entrada conteiner**:
  - `CNTR`: mascara curta (prefixo visual de 5 caracteres do final).
  - `MskCntr`: mascara do ID completo ISO-like.
- Botao **X** (`Command4`) no canto superior direito — fecha.

### Faixa de setpoint (somente leitura)

- Labels **MIN** / **MAX** (`Label8`, `Label10`) + campos `txtMIN`, `txtMAX` (**somente ReeferS**).
- **SET POINT** (`Label9`) + `txtSetPoint`.
- Coluna esquerda adicional: **UMIDADE (%)** setpoint (`txtUmidadeSet`), **VENTILACAO m3/h** setpoint (`txtVentilacaoSet`).

### Leituras (entrada)

- **TEMP SUP.**, **TEMP RET.**, **ESCALA** (`txtTempSup`, `txtTempRet`, `txtEscala`) na mesma faixa.
- **UMIDADE (%)** e **VENTILACAO** medidos (`txtUmidade`, `txtVentilacao`).
- **GATE IN** (`Label7` + `txtGateIN` locked) — informacao de entrada.

### Acoes

- **PLUG OFF** (`Command2`) — botao de acao critica.
- **SALVAR** (`Command1`).

### Grid historico

- `Grid2` — linhas de historico com colunas: Data, Temp SUP, Temp RET, Umidade, Ventilacao, Usuario (formato definido em `FormatString`).

## Aba "Previsao"

### Bloco entradas

- Titulo funcional: **"UNIDADES PREVISTAS - ENTRADA PATIO"** (`Label1(1)`).
- Contador `lblTotEnt` (fundo preto, numero laranja).
- `GridEntradas` — colunas: Conteiner, Tam, Set Point, Atracacao, Local Atracacao, Entrada Prevista.

### Bloco saidas / estoque

- Titulo **"UNIDADES EM ESTOQUE"** (`Label1(2)`).
- Contador `lblTotSai`.
- Tres **checkboxes**:
  - "Apenas os Desligados"
  - "Com agendamento Saida"
  - "Posicionados"
- `GridSaidas` — colunas de gate, yard, previsao saida, setpoint, primeira e ultima medicao; **ultima coluna oculta** (`COLWIDTH` = 0) usada para armazenar `dt_plug_off` e pintar linha (`CellBackColor` = cor de `Shape1`) quando plug off existir.

## Aba "Un. Desligadas"

- `Label13` (titulo auxiliar — caption vazio no designer, pode ser preenchido em runtime em outro trecho nao visto; **Hipotese**: titulo estatico em outra versao ou permanece vazio).
- `GridDesligados` com coluna extra **Data Plug OFF** vs grid de saidas.

## Atalhos / teclado

| Tecla | Comportamento declarado |
|-------|-------------------------|
| ESC | Fecha form |
| Enter | `SendKeys "{tab}"` em `KeyDown` — ver bug de parametro no diagnostico |

## Equivalencia web sugerida (Razor / HTML / JS)

| VB6 | Web sugerido |
|-----|----------------|
| `SSTab` | `nav-tabs` + paineis Bootstrap (ou tabs proprias ja usadas no Romaneio) |
| `MaskEdBox` | `<input>` com mascara JS (ex.: IMask) + validacao de tamanho |
| `TextBox Locked` | `readonly` + estilo visual diferenciado |
| `CommandButton` | `<button type="button">` primario/ secundario / perigo para PLUG OFF |
| `MSHFlexGrid` | `<table class="table">` com DataTables simples ou grid ja padrao do projeto |
| Contadores `lblTot*` | badge ou texto ao lado do titulo da secao |
| Checkbox filtros | mesmos filtros como querystring ou POST ao recarregar sub-grid |

## UX web — adaptacoes recomendadas

1. **Separar acoes destrutivas**: `PLUG OFF` deve ter **confirmacao modal** explicita (no VB6 so guarda por `lblReeferDesligado.Visible`).
2. **Feedback de sucesso/erro**: padrao JSON + toast/alerts como demais telas Romaneio (`JsonResult` com `success`/`message`).
3. **Enter para proximo campo**: implementar com JS no container do formulario, sem `SendKeys`.
4. **Responsividade**: layout atual e fixo em twips; migrar para grid CSS ou cards empilhaveis em mobile.
5. **MIN/MAX**: exibir proximo ao setpoint como **contexto** para validacao (hoje nao ha validacao automatica contra min/max no VB6 analisado).

## Checklist interno (Agente 03)

- [x] Controles inventariados
- [x] Equivalente web proposto
- [x] Estados dinamicos (alertas, highlight de linhas) mapeados
- [x] Atalhos / teclado
- [x] Riscos de usabilidade documentados
