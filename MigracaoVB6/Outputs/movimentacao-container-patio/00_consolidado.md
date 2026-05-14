# 00 — Consolidado — Movimentação Container: Pátio

## Resumo executivo

A funcionalidade **“Movimentação Container: Pátio”** corresponde, no legado VB6, ao form de classe **`FrmMovimentacao`**, aberto pelo menu **“Movimentação Conteiner”** (`PrincipalS.frm` índice 4 no pacote **Coletor Pátio**). O código-fonte principal para o cenário de **pátio** está em **`FrmMovimentacaoS.frm`**, compilado pelo projeto **`ColetorPatioSmall.vbp`** (“Coletor Pátio”).

A tela permite **consultar** o contêiner no estoque do pátio, **visualizar** alertas operacionais (reefer, bloqueios, SIGVIG, carregamento registrado, etc.), **planejar** destino e motivo, e **executar** a movimentação chamando **`operador.dbo.fnMovimentaCntr`** (pré-checagem) e **`operador.dbo.pcMovimentaCntr`** (persistência). Há fluxos especiais (**SAÍDA**, **DTA**, **BK**, **CLIA**, **IMO** com possível estorno), validações de **lacres** e modal **`FrmVeiculoServ`**. A variante **S** inclui gestão de **avarias** do contêiner com permissões **`TB_SYS_FUNCOES`** (códigos **998** e **999**).

---

## Decisões confirmadas (com evidência)

1. **Nome de menu vs nome de produto:** o menu usa “Movimentação Conteiner”; o executável de pátio é “Coletor Pátio” — ambos apontam para o mesmo `VB_Name` (`FrmMovimentacao`).
2. **Núcleo de persistência:** `pcMovimentaCntr` após `fnMovimentaCntr` — evidência direta em `Atualiza_Posicao`.
3. **Filtro de pátio:** uso intensivo de `Col_Patio` com **`IN (1,7)`** quando o pátio da sessão é 1 ou 7; lógica especial para **`Col_Patio = 3`** em `tb_yard` / substring de pilha.

---

## Lista de arquivos VB6 relacionados

| Arquivo | Papel |
|---------|--------|
| `ProjetoReferencia/Coletor/FrmMovimentacaoS.frm` (+ `.frx`) | Variante **Pátio Small** — base principal deste estudo |
| `ProjetoReferencia/Coletor/FrmMovimentacao.frm` (+ `.frx`) | Variante alternativa (layout diferente) |
| `ProjetoReferencia/Coletor/PrincipalS.frm` | Abre `FrmMovimentacao` (menu Pátio) |
| `ProjetoReferencia/Coletor/Principal.frm`, `PrincipalG.frm` | Mesma chamada em outros menus |
| `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp` | Projeto “Coletor Pátio” |
| `ProjetoReferencia/Coletor/Frmmensagem.frm`, `frmMensagem2.frm` | Mensagens |
| `ProjetoReferencia/Coletor/FrmVeiculoServ.frm` | Modal SAÍDA/DTA |
| `ProjetoReferencia/Coletor/Microled.bas`, `mdlColetor.bas`, `Segreg_IMO.bas` | Infra e regra IMO |

---

## Artefatos gerados (esta pasta)

| Arquivo | Conteúdo |
|---------|----------|
| `01_diagnostico_vb6.md` | Eventos, SQL, dependências, riscos técnicos |
| `02_mapeamento_funcional.md` | Fluxos, dados lidos/gravados, impacto operacional |
| `03_mapeamento_visual.md` | Controles → Razor/HTML/JS |
| `04_regras_negocio.md` | Regras classificadas (RN/RI/RT/RV/H) |
| `05_desenho_mvc.md` | Proposta aderente ao Romaneio |
| `06_plano_implementacao.md` | Fases, DoD, dependências |

---

## Principais regras identificadas

1. **Motivo obrigatório** e **destino obrigatório/diferente da origem**.
2. **Reefer sem temperatura** pode bloquear movimentação.
3. **Lacres:** várias validações sequenciais + checkbox **“LACRES CONFERIDOS”** quando histórico exige conferência.
4. **SAÍDA:** confirmação + **`FN_BLOQUEIO_SCANNER`** para importação.
5. **Motivos com câmera:** seleção obrigatória de **`TB_CAMERAS`** antes da SP.
6. **BK/CLIA:** flags na **`TB_CNTR_BL`** + regras de reserva/agenda.
7. **IMO:** possível **estorno** via segunda execução da SP invertendo yards.
8. **Avarias:** CRUD restrito por **finalização** (`TB_ETAPAS_AVARIA`) e **permissões**.

---

## Objetos de banco envolvidos (alto nível)

- **TVFs:** `fnInvent_Sistemas_YardF`, `fnInvent_Sistemas_Col`
- **Funções:** `fnMovimentaCntr`, `FN_BLOQUEIO_SCANNER`
- **SP:** `pcMovimentaCntr`
- **Views:** família `VW_PROX_MOV_*`, `VW_EGMCI_AVARIAS_TTL`, `VW_AGG_AGENDAMENTOS`, `VW_SIGVIG_CNTR_ESTOQUE`, etc.
- **Tabelas:** `TB_CNTR_BL`, `TB_HIST_SHIFTING`, `TB_YARD`, `TB_cad_motivo`, `TB_CAMERAS`, `TB_PATIOS`, `TB_AVARIAS_CONTEINER`, `TB_ETAPAS_AVARIA`, `dte_tb_avarias`, cadastros SIGVIG (lista completa no `01`)

---

## Proposta inicial de migração MVC

- **`MovimentacaoContainerPatioController`** + **`IMovimentacaoContainerPatioRepositorio`** seguindo o padrão JSON de `LocalizacaoCarregamentosController`.
- **View única** com abas (Bootstrap `nav-tabs`) espelhando `SSTab1`.
- **JavaScript** dedicado para busca, grids e fluxo de câmera em dois passos.
- **Repositório** centraliza validações e chamadas às rotinas de banco já existentes (sem reimplementar negócio SQL nas controllers).

Detalhes em `05_desenho_mvc.md`.

---

## Critérios de aceite (migrado = pronto para operação)

### Consulta e filtros

- [ ] Busca por **final** replica mensagens de **nenhum** / **mais de um** contêiner.
- [ ] Busca por ID completa carrega dados apenas se unidade existir em `fnInvent_Sistemas_Col` para o **pátio da sessão** (regra 1/7/3 preservada).

### Exibição

- [ ] Yard atual, destino, tipo/tamanho, navio/viagem, lacres, IMO, GMCI, SIGVIG (quando aplicável), alertas de bloqueio e “carregamento registrado” aparecem conforme regras do legado.
- [ ] Grids de **próximos movimentos** e **vizinhança de pilha** refletem os mesmos filtros que o VB6 para cenários de teste acordados.

### Ações

- [ ] **Salvar** bloqueia com as mesmas mensagens nas mesmas ordens do VB6.
- [ ] Motivos com **flag_camera** abrem fluxo de câmera antes da SP.
- [ ] **SAÍDA** exige confirmação e respeita bloqueio de scanner.
- [ ] Fluxos **BK** e **CLIA** aplicam as mesmas validações de flags e agenda observadas no código.
- [ ] **DTA** e **SAÍDA com EF=E** acionam o fluxo equivalente ao modal `FrmVeiculoServ` (paridade acordada com negócio).

### Pós-movimentação

- [ ] Unidade altera posição de forma consistente com o inventário (validação cruzada em consulta já existente ou script).
- [ ] Divergência de lacres gera avaria automática quando aplicável.
- [ ] Cenário IMO dispara alerta e **não deixa estado inconsistente** (estorno ou transação única aprovada).

### Avarias (variante S)

- [ ] Permissões 998/999 reproduzem habilitação da aba e botões.
- [ ] Finalização impede novas edições como no legado.

---

## Riscos e dúvidas abertas

| Item | Tipo |
|------|------|
| Corpo de `pcMovimentaCntr` / transações internas não versionados | Risco técnico |
| SQL com sintaxe Oracle em `Busca_Dados` | Risco de ambiente / lacuna |
| Comportamento completo de `FrmVeiculoServ` | Lacuna funcional |
| `Coletor.vbp` sem referência explícita ao form | Lacuna de build |
| Dupla chamada à SP para IMO | Risco de consistência |
| `Form_Unload` recursivo | Risco legado (não replicar no web) |
| Volume elevado de SQL concatenado no VB6 | Risco de segurança na migração |

---

## Status de prontidão para implementação

| Área | Status |
|------|--------|
| Localização no legado | **Pronta** |
| Mapeamento funcional/técnico | **Pronta** (com hipóteses marcadas) |
| Desenho MVC | **Pronta** (proposta) |
| Plano incremental | **Pronta** |
| DDL/procedures no repositório | **Não disponível aqui** — validação obrigatória em ambiente SQL antes da Fase 4 |
| Paridade total sem homologação operacional | **Não pronta** |

**Conclusão:** o estudo está **adequado para iniciar o Agente 07 (implementação MVC)** após validação rápida das rotinas de banco e decisão sobre o fluxo `FrmVeiculoServ` e a query de agendamento Oracle.

---

## Próximo passo sugerido

Executar o prompt do **`MigracaoVB6/Agents/07_agente_execucao_mvc.md`** usando esta pasta como insumo (`01`–`06` + este consolidado).
