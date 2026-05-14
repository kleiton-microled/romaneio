# 06 — Plano de implementação — Movimentação Container: Pátio

## Visão geral

Entrega incremental em camadas: primeiro **leitura/validação** (baixo risco), depois **movimentação** (alto risco), por último **avarias** e **polimento UX**. Cada etapa possui critérios de pronto testáveis contra o comportamento descrito nos artefatos 01–04.

---

## Fase 0 — Preparação e contratos

**Objetivos**

- Congelar lista de endpoints e DTOs (documento 05).
- Validar no banco: assinaturas de `fnInvent_Sistemas_YardF`, `fnInvent_Sistemas_Col`, `fnMovimentaCntr`, `pcMovimentaCntr`, `FN_BLOQUEIO_SCANNER`.
- Resolver destino do SQL “Oracle” embutido em `Busca_Dados` (remover, portar ou isolar).

**Pronto quando**

- Script ou documento interno confirma parâmetros das rotinas e exemplos de retorno OK/erro de `fnMovimentaCntr`.
- Decisão registrada sobre agendamento de posicionamento.

**Riscos:** dependência de objetos não versionados no git; necessidade de ambiente SQL real.

---

## Fase 1 — Shell MVC + sessão + página

**Objetivos**

- Controller + View + rota principal.
- Carregar descrição do pátio (`TB_PATIOS`) e motivos (`TB_cad_motivo`), câmeras (`TB_CAMERAS`) como no `Form_Load`.

**Pronto quando**

- Usuário logado abre a página e vê pátio correto conforme `Session["Patio"]`.
- Motivos listados com Select2 (ou equivalente já adotado).

**Testes:** smoke manual + unauthorized redirect.

---

## Fase 2 — Busca por final e carregamento de detalhe

**Objetivos**

- Implementar `BuscarPorFinal` e `CarregarDetalhe` espelhando `Busca_Cntr` / `Busca_Dados`.
- Exibir painel de dados e badges (Reefer, DTA, desova, alertas de bloqueio conforme queries).

**Pronto quando**

- Mensagens de zero / múltiplos resultados coincidem com o legado.
- `Autonum` do contêiner fica disponível para próximas chamadas.

**Testes:** casos com pátio 1, 7 (IN), pátio 3 (yard_eq), sistemas I/A/R se disponíveis em QA.

---

## Fase 3 — Grids de apoio

**Objetivos**

- `ProximosMovimentos` (views `VW_PROX_MOV_CNTR*`).
- `VizinhancaPilha` reproduzindo ramificações de `Carrega_Grid2` (inclui lógica par/impar).

**Pronto quando**

- Comparativo lado a lado VB6 vs Web em ambiente de homologação com mesmo dataset mostra mesmas linhas (ordenação incluída).

**Riscos:** regressão em filtros LIKE — priorizar testes com yards reais.

---

## Fase 4 — Movimentação (núcleo)

**Objetivos**

- Portar validações de `Atualiza_Posicao` na mesma ordem.
- Invocar `fnMovimentaCntr` → `pcMovimentaCntr`.
- Fluxo de câmera em dois passos.
- SAÍDA + scanner + Reefer + lacres + BK/CLIA + DTA/SAÍDA com `FrmVeiculoServ` (wrapper web).

**Pronto quando**

- Movimentação bem-sucedida atualiza yard no inventário e registra histórico como no legado (validar via consultas diretas ou tela de consulta existente).
- Mensagens de erro da função aparecem ao usuário sem expor stack trace.

**Riscos:** alto — exige homologação operacional; recomenda-se feature flag ou piloto em pátio não crítico.

---

## Fase 5 — Divergência de lacres e segregação IMO

**Objetivos**

- Repetir lógica `FlagDivLacre` e insert condicional em `TB_AVARIAS_CONTEINER`.
- Portar `Verifica_Regras_Seg_Imo_Delta` + estorno via segunda SP.

**Pronto quando**

- Casos de teste com IMO forçado reproduzem alerta e posição volta ao estado anterior.

**Riscos:** segundo call à SP pode falhar se estado intermediário for inválido — ter plano B (transação única no backend).

---

## Fase 6 — Aba Avarias

**Objetivos**

- Permissões espelhando funções 998/999.
- Incluir / excluir / finalizar com mesmas restrições (`Verifica_Finalizado`).
- Carregar `VW_EGMCI_AVARIAS_TTL` e histórico visual.

**Pronto quando**

- Usuário sem permissão não vê ações.
- Fluxos de bloqueio por finalização funcionam.

---

## Fase 7 — UX, telemetria e hardening

**Objetivos**

- Responsividade, dupla confirmação de SAÍDA, prevenção de double-click.
- Parametrização total dos SQL novos.
- Logs mínimos de auditoria (se aprovado).

**Pronto quando**

- Checklist de segurança básica (sem concatenação aberta, timeout configurável).

---

## Dependências entre fases

```
0 → 1 → 2 → 3 → 4 → 5
              ↘ 6 (pode iniciar após 2 se usar apenas listagens; movimentação final precisa da Fase 4 estável)
7 → encerra após 4–6
```

---

## Checklist técnico final (DoD macro)

- [ ] Paridade de mensagens críticas validada com operação
- [ ] Pátios 1/7/3 cobertos em testes
- [ ] Motivo+câmera funcionando
- [ ] SAÍDA + scanner + Reefer
- [ ] BK/CLIA validados com usuários de negócio
- [ ] IMO / rollback validado
- [ ] Avarias com perfis
- [ ] Sem SQL concatenado na nova camada
- [ ] Documentação de rotinas atualizada no wiki interno (se existir)

---

## Estimativa de esforço (ordem de grandeza, não compromisso)

| Fase | Complexidade |
|------|----------------|
| 0–2 | Média |
| 3 | Alta |
| 4 | Muito alta |
| 5–6 | Alta |
| 7 | Média |

---

## Itens que exigem validação com usuário final

- Aceitação de substituir dupla chamada `pcMovimentaCntr` por transação única na migração IMO.
- Comportamento do fluxo **FrmVeiculoServ** (campos obrigatórios).
- Prioridade da query Oracle/agendamento — ainda é usada?
