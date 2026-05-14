# Consolidado — Container REEFER (Patio)

## Resumo executivo

A funcionalidade **"Conteiner - REEFER"** do **Coletor Pátio** corresponde ao form VB6 `Reefer` implementado em `ReeferS.frm`, aberto pelo menu `PrincipalS.frm` (`Command1` index **7**). A tela combina **monitoramento operacional de reefer** (gravacao em `OPERADOR..TB_MONITORING`), **gestao de plug** (`DT_PLUG_OFF` em `SGIPA..TB_CNTR_BL` e, apos salvar, tambem em `REDEX..TB_PATIO` no fluxo Patio), e **visoes de previsao/listagem** de conteneres reefer por patio. O contexto **Pátio** esta evidenciado pelo `ColetorPatioSmall.vbp` (`Title="Coletor Pátio"`) e pelos filtros `PATIO IN (1,7)` quando `Col_Patio` e 1 ou 7.

Existe variante paralela `Reefer.frm` (coletor armazem) com **diferencas relevantes** de filtro SQL, gravacao de coluna `PATIO` e ausencia de update em `REDEX..TB_PATIO` — nao e o foco principal deste estudo, mas consta como referencia cruzada.

## Decisoes confirmadas (com base em codigo)

1. **Fonte principal Patio**: `ProjetoReferencia/Coletor/ReeferS.frm` + `ReeferS.frx`.
2. **Entrada de menu**: `PrincipalS.frm` — `Case 7: Reefer.Show 1`.
3. **Projeto executavel Patio**: `ColetorPatioSmall.vbp` referencia `ReeferS.frm`.
4. **Persistencia central**: `OPERADOR..TB_MONITORING` para historico; `SGIPA..TB_CNTR_BL` para plug; `REDEX..TB_PATIO` para limpar plug apos salvar (Patio).
5. **Inventario**: consulta a `<Banco_Operador>VW_INVENT_SISTEMAS` com filtro de patio dinamico.

## Hipoteses em aberto

- Texto exato dos forms de mensagem compilados no exe Patio (`Frmmensagem` vs `FrmmensagemS`).
- Origem e garantias de `Cod_Usuario` global usado no INSERT.
- Semantica real do SQL `CASE V.DT_ATRACACAO WHEN NULL` em producao.
- Impactos indiretos sobre **romaneio** ou **carregamento** via outros processos nao mapeados neste form.

## Itens bloqueantes antes de codificar

1. Validacao de negocio sobre **correcao do bug** `txtUmidadeSet` recebendo `"CLOSED"` (ventilacao vazia) — migrar fielmente ou corrigir.
2. Confirmacao DBA sobre uso de `SUBSTR(ID_CONTEINER,8,5)` com o formato real de `ID_CONTEINER` no ambiente alvo.
3. Definir se **PLUG OFF** deve ser estendido a sistemas **R/O** (hoje so IPA no `Command2`).

## Escopo da primeira implementacao sugerida

- Aba **Monitoramento** completa (busca, detalhe, historico, salvar, plug off).
- Aba **Previsao** e **Un. Desligadas** em segunda entrega **ou** junto se prazo permitir (SQL ja levantado).

## Criterios de aceite (migracao considerada adequada)

### Funcionais

- [ ] Busca por **final** replica contagem, unicidade e mensagens equivalentes ao legado (ReeferS).
- [ ] Busca por **ID completo** aplica filtros de patio **incluindo unificacao 1/7** quando aplicavel.
- [ ] Bloqueios funcionais: nao reefer; reefer "DRY"; conteiner nao encontrado.
- [ ] Campos obrigatorios na gravacao espelham `Valida_Dados`.
- [ ] Alerta de divergencia **supply vs setpoint** > 2 com possibilidade de cancelar operacao.
- [ ] **Salvar** grava `TB_MONITORING` com usuario e patio (`Col_Patio=7` gravado como **1** no legado — preservar se confirmado).
- [ ] **Salvar** limpa `DT_PLUG_OFF` em `TB_CNTR_BL` (IPA) e `TB_PATIO` (RDX) conforme ReeferS.
- [ ] **PLUG OFF** grava `DT_PLUG_OFF=GETDATE()` apenas quando guardas equivalentes forem satisfeitas.
- [ ] Historico ordenado como no legado (`AUTONUM DESC`).
- [ ] Abas de listagem retornam colunas equivalentes e contadores (`RecordCount`).

### Filtros e dados exibidos

- [ ] `PATIO` da sessao Romaneio aplicado com mesma regra `IN (1,7)` quando `Col_Patio` legado seria 1 ou 7.
- [ ] Setpoints, Gate IN, MIN/MAX (Patio), alertas de plug off visiveis como no legado.

### Regras de reefer / energia / temperatura a preservar

- [ ] Flags `FLAG_REEFER` / `FLAG_REEFER_DESLIGADO` na resolucao por final (ReeferS).
- [ ] Ciclo **PLUG OFF manual** vs **limpeza ao salvar leitura**.
- [ ] Historico de medidas por conteiner (chaves `AUTONUM_CNTR_*`).

### Comportamentos VB6 a manter vs adaptar

| Comportamento | Manter | Adaptar |
|---------------|--------|---------|
| Concatenacao SQL | | X (parametrizar) |
| Ausencia de transacao | | X (transacao no salvar) |
| `SendKeys` Enter | | X |
| Bug ventilacao/umidade setpoint | HP | Corrigir se aprovado |
| Highlight cor em grid | | X (flag no JSON + CSS) |

## Objetos de banco envolvidos (lista curta)

- **Tabelas**: `OPERADOR..TB_MONITORING`, `SGIPA..TB_CNTR_BL`, `REDEX..TB_PATIO`, `SGIPA..TB_CAD_USUARIOS`, `SGIPA..DTE_TB_VIAGENS`, `SGIPA..DTE_TB_ARMAZENS`, `SGIPA..TB_TTIMEUNT`, `OPERADOR..TB_GD_RESERVA`
- **View**: `<Banco_Operador>VW_INVENT_SISTEMAS`
- **Procedures**: nenhuma no legado desta tela

## Riscos e lacunas

| Item | Tipo |
|------|------|
| SQL injection no legado — nao replicar | Risco tecnico |
| Sem transacao no multi-update | Risco de consistencia |
| `Form_KeyDown` com `KeyAscii` incorreto | Lacuna de UX / bug legado |
| `CASE WHEN NULL` em SQL | Lacuna de resultado |
| Comportamento multi-sistema R/O para PLUG OFF | Lacuna funcional |
| Impacto romaneio/carregamento | Hipotese |

## Arquivos VB6 relacionados (lista)

- `ProjetoReferencia/Coletor/ReeferS.frm`, `ReeferS.frx`
- `ProjetoReferencia/Coletor/Reefer.frm`, `Reefer.frx` (variante armazem)
- `ProjetoReferencia/Coletor/PrincipalS.frm`
- `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp`
- `ProjetoReferencia/Coletor/Microled.bas`
- `ProjetoReferencia/Coletor/mdlColetor.bas`
- Forms de mensagem: `FrmmensagemS.frm` / `Frmmensagem.frm`, `frmMensagem2S.frm` / `frmMensagem2.frm` (confirmar paridade build)

## Artefatos gerados (pasta desta migracao)

- `MigracaoVB6/Outputs/container-reefer-patio/01_diagnostico_vb6.md`
- `MigracaoVB6/Outputs/container-reefer-patio/02_mapeamento_funcional.md`
- `MigracaoVB6/Outputs/container-reefer-patio/03_mapeamento_visual.md`
- `MigracaoVB6/Outputs/container-reefer-patio/04_regras_negocio.md`
- `MigracaoVB6/Outputs/container-reefer-patio/05_desenho_mvc.md`
- `MigracaoVB6/Outputs/container-reefer-patio/06_plano_implementacao.md`
- `MigracaoVB6/Outputs/container-reefer-patio/00_consolidado.md`

## Proposta inicial de migracao MVC (recapitulacao)

- `ContainerReeferPatioController` + `ContainerReeferPatioRepositorio` (Dapper) + ViewModel + `Index.cshtml` + JS dedicado.
- Reutilizar `MontarFiltroPatio` e padroes de sessao como `MovimentacaoContainerPatioController`.

## Status de prontidao para implementacao

| Criterio | Status |
|----------|--------|
| Localizacao legado | **Pronto** |
| SQLs e fluxos mapeados | **Pronto** |
| Divergencias Reefer vs ReeferS documentadas | **Pronto** |
| Imagem de referencia no repo | **Ausente** (nao obrigatoria para continuar) |
| Decisoes de negocio sobre bug e PLUG OFF R/O | **Pendente** |
| Validacao DBA de views/filtros | **Recomendado** |

**Conclusao**: estudo **pronto para iniciar implementacao** apos fechamento dos itens bloqueantes de negocio/DBA listados acima.
