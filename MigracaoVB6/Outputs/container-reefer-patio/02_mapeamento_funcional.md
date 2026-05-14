# Mapeamento funcional — Container REEFER (Patio)

## Finalidade da tela (negocio)

Permitir ao operador de **patio**:

1. **Registrar leituras periodicas de reefer** (temperatura supply/retorno, umidade, ventilacao, escala) para um **contener reefer** localizado no patio corrente (`Col_Patio`), persistindo historico em `OPERADOR..TB_MONITORING` e vinculando o registro ao contener nos sistemas IPA / Redex / Operador conforme `VW_INVENT_SISTEMAS.SISTEMA`.
2. **Consultar previsoes e listagens** de conteneres reefer: entradas previstas no patio, unidades em estoque (com filtros), e unidades com plug desligado.
3. **Registrar PLUG OFF** (energia) em `SGIPA..TB_CNTR_BL.DT_PLUG_OFF` quando aplicavel ao ramo IPA.
4. Ao **salvar leitura com sucesso**, **limpar** `DT_PLUG_OFF` (reativar registro energetico) em `TB_CNTR_BL` e, no fluxo Patio (`ReeferS`), tambem em `REDEX..TB_PATIO` quando o contener estiver no sistema Redex.

## Pre-condicoes

- Coletor Patio autenticado com `Col_Patio` definido (variavel global do coletor).
- Conexao `db` ativa (`Microled.bas`).
- Usuario com codigo numerico em `Cod_Usuario` para gravacao em `TB_MONITORING.USUARIO` (**Hipotese**: preenchido no login — nao declarado neste `.frm`).

## Fluxo principal (feliz) — Monitoramento

1. Operador informa **final do conteiner** em `CNTR` (5 caracteres) OU ID completo em `MskCntr`.
2. Se final: `Busca_Cntr` valida existencia unica em `VW_INVENT_SISTEMAS` (reefer, nao desligado, patio filtrado).
3. Ao completar/sair do `MskCntr`, `Busca_Dados` carrega:
   - setpoint, umidade/vent setpoint, escala, gate in;
   - alertas **reefer desligado** / **plug off** quando `DT_PLUG_OFF` preenchido (IPA: `TB_CNTR_BL`);
   - `txtMIN`/`txtMAX` (**apenas ReeferS**) para faixa declarada no cadastro do BL;
   - historico `Grid2`.
4. Operador preenche **leituras atuais** (SUP, RET, UMIDADE, VENTILACAO, ESCALA).
5. `SALVAR`: validacoes; INSERT em `TB_MONITORING`; opcionalmente zera `DT_PLUG_OFF`; limpa leituras e retorna foco ao final do conteiner.

## Fluxos alternativos / erro

| Condicao | Comportamento evidenciado |
|----------|---------------------------|
| Final sem conteiner | Mensagem "Nenhum conteiner Reefer com este final" (ReeferS) / variante sem "Reefer" no texto em `Reefer.frm` |
| Mais de um conteiner com mesmo final | Mensagem pedindo identificacao completa |
| Conteiner nao encontrado na view | "Conteiner nao encontrado" |
| `flag_reefer = N` | "Conteiner nao e reefer" |
| `flag_reefer_desligado = Y` | "Conteiner reefer captado com carga DRY!" |
| Validacao campos vazios | Mensagens especificas por campo |
| Supply divergente >2 do setpoint | `frmMensagem2` com opcao de cancelar |
| PLUG OFF com alerta "REEFER DESLIGADO" visivel | Botao `Command2` **nao** executa update (guarda de UI) |

## Acoes do usuario

| Acao | Tipo |
|------|------|
| Informar final / ID conteiner | Entrada |
| Alternar abas Previsao / Desligadas | Navegacao |
| Marcar filtros nas saidas | Filtro (`Check1..3`) |
| Salvar monitoramento | Operacao de escrita |
| PLUG OFF | Operacao de escrita (data plug) |
| Fechar tela (`X` ou ESC) | Sair |

## Dados consultados (visao funcional)

- Inventario multi-sistema: `VW_INVENT_SISTEMAS` (patio, flags reefer, sistema I/R/O).
- Cadastro BL (IPA): `TB_CNTR_BL` (datas plug off, temp min/max — ReeferS).
- Historico: `TB_MONITORING` + usuario.
- Previsao entrada: join `TB_CNTR_BL` + viagens + armazens + `TB_TTIMEUNT`.
- Estoque / desligados: `TB_CNTR_BL` + agregados min/max de `TB_MONITORING` + `TB_GD_RESERVA`.

## Dados gravados / alterados

| Destino | Operacao |
|---------|----------|
| `OPERADOR..TB_MONITORING` | INSERT cada salvar |
| `SGIPA..TB_CNTR_BL.DT_PLUG_OFF` | UPDATE GETDATE() (plug off); UPDATE NULL (apos salvar leitura se IPA) |
| `REDEX..TB_PATIO.DT_PLUG_OFF` | UPDATE NULL apos salvar leitura se RDX (**ReeferS apenas**) |

## Impacto em patio / conteiner / carregamento

- **Patio**: filtro constante `Col_Patio`; em Patio, **unifica patios 1 e 7** nas consultas quando `Col_Patio` e 1 ou 7 (evidencia em `ReeferS`).
- **Reefer / energia**: altera `DT_PLUG_OFF` (status plug).
- **Temperatura operacional**: historiza leituras; alerta divergencia vs setpoint.
- **Romaneio / carregamento**: **sem referencia direta** neste `.frm` — impacto apenas **Hipotese** via consistencia de dados compartilhados (`TB_CNTR_BL`, patio).

## Mensagens (texto funcional — corrigir encoding na migracao)

- "Nenhum conteiner Reefer com este final" / "Nenhum conteiner com este final"
- "Existe mais de 1 conteiner reefer com este final..." / "...mais de 1 conteiner..."
- "Conteiner nao encontrado"
- "Conteiner nao e reefer"
- "Conteiner reefer captado com carga DRY!"
- "Conteiner invalido"
- "Temperatura fornecida nao informada" / retorno / umidade / ventilacao / escala
- `Col_Mensagem`: "Atencao. Temperatura fornecida(Supply) muito divergente da temperatura de Set Point"

## Consulta vs operacao

A tela **nao e somente consulta**: executa **INSERT** de monitoramento, **UPDATE** de plug off (manual e reversao ao salvar), e **consultas** de apoio nas abas.

## Paridade Reefer.frm (Armazem) — resumo

| Aspecto | ReeferS (Patio) | Reefer (Arm) |
|---------|-----------------|--------------|
| Filtro `VW_INVENT_SISTEMAS` patio | `IN (1,7)` se `Col_Patio` 1 ou 7 | `= Col_Patio` |
| Flags reefer na primeira query | Ativos no SQL | Comentados (nao aplicados na string ativa) |
| INSERT coluna `PATIO` | `IIf(Col_Patio=7,1,Col_Patio)` | `Col_Patio` |
| Pos-salvar limpar plug Redex | Sim (`TB_PATIO`) | Nao |
| Campos MIN/MAX temperatura | Sim | Nao |

## Checklist interno (Agente 02)

- [x] Finalidade descrita em negocio
- [x] Acoes principais e auxiliares
- [x] Fluxo feliz
- [x] Fluxos alternativos/erro
- [x] Campos obrigatorios (via `Valida_Dados` + uso operacional)
- [x] Mensagens relevantes listadas
