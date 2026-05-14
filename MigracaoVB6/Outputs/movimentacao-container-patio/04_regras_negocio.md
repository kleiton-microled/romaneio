# 04 — Regras de negócio — Movimentação Container: Pátio

Legenda: **[RN]** Negócio · **[RI]** Interface · **[RT]** Técnica · **[RV]** Validação · **[H]** Hipótese

---

## Por fase: identificação e carga

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R01 | [RV] | Final do ID deve resolver para **exatamente um** contêiner no pátio; zero ou >1 bloqueiam | `Busca_Cntr` + `fnInvent_Sistemas_YardF` |
| R02 | [RN] | Estoque visível é sempre filtrado pelo **`Col_Patio`** (com `IN (1,7)` quando pátio 1 ou 7) | Várias queries |
| R03 | [RV] | Se `fnInvent_Sistemas_Col` não retornar linha, abortar com “Unidade não encontrada no estoque” | `Busca_Dados` |
| R04 | [RN] | O “sistema” da linha (`I`, `A`, `R`) determina mensagens, SQL de histórico e parte das validações | `MskCntr.Tag` |
| R05 | [RT] | Campos `.Tag` dos lacres guardam códigos do manifesto para comparar divergência | Atribuições em `Busca_Dados` |
| R06 | [H] | Query com sintaxe Oracle para agendamentos de posicionamento pode não rodar em todos os ambientes | Trecho `tb_agendamento_posicao` |

---

## Motivo, destino e UI auxiliar

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R07 | [RV] | **Motivo** obrigatório antes de salvar | `Command1_Click`, `Atualiza_Posicao` |
| R08 | [RV] | **Destino** obrigatório e **≠** yard atual | `Atualiza_Posicao` |
| R09 | [RI] | Alternar opção ATUAL/DESTINO troca label “FIADA” e recarrega `Grid2` | `Option1_Click`, `Carrega_Grid2` |
| R10 | [RI] | Ao mudar destino/origem, pode pré-selecionar motivo (`DataCombo2.BoundText` 1, 8, 14…) | `mskDestino_LostFocus`, `TXTYardAtual_LostFocus` |
| R11 | [RN] | Se motivo possui **`flag_camera=1`**, movimentação só prossegue após escolher **câmera** válida | `Command1_Click`, `Command3_Click` |

---

## Reefer, lacres e checkpoints operacionais

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R12 | [RN]+[RV] | Reefer com **`flag_reefer_sem_temp=1`** e **não desligado** → bloqueio de posicionamento | `Atualiza_Posicao`, também alerta pós-`Busca_Dados` |
| R13 | [RN]+[RV] | Para sistema **`I`**, se não existe histórico shifting com origem ≠ GATEIN, operador deve marcar **LACRES CONFERIDOS** | Query `TB_HIST_SHIFTING` + `Check1` |
| R14 | [RN]+[RV] | Para **`I`**, exige pelo menos um lacre de origem entre IPA/LACRE1-3 | `Atualiza_Posicao` |
| R15 | [RN]+[RV] | Para **`I`**, **`txtEF=F`** e tipo **`OT`**, lacre lona obrigatório | `Atualiza_Posicao` |
| R16 | [RV] | Preenchimento sequencial de lacres Gate (não pular “buracos”) exceto quando destino é **DTA** | Bloco “Lacres informados incorretamente” |

---

## Saída e bloqueios especiais

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R17 | [RN]+[RV] | Destino **`SAIDA`** exige confirmação explícita | `MsgBox` |
| R18 | [RN]+[RV] | Na saída, sistema **`I`** passa por **`FN_BLOQUEIO_SCANNER`** — bloqueio impede operação | `Atualiza_Posicao` |
| R19 | [RN] | **`EF=E`** + **`SAIDA`** abre **`FrmVeiculoServ`** antes da movimentação | `Atualiza_Posicao` |
| R20 | [RN] | Fluxo **DTA** abre **`FrmVeiculoServ`** e redefine destino textual para **CAM** após retorno | `Atualiza_Posicao` |

---

## Destinos BK / CLIA

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R21 | [RN]+[RV] | **BK** só permitido se `FLAG_SEL_TRANSF_BK` na unidade | `TB_CNTR_BL` |
| R22 | [RN] | **BK** pode reposicionar reserva (`AUTONUM_GD_RESERVA`) conforme janela de movimentos vs limite | Updates/SELECTs `TB_GD_RESERVA` |
| R23 | [RN]+[RV] | Se existir agendamento conflitante com janela BK, bloquear com mensagem detalhada | `VW_AGG_AGENDAMENTOS` + texto auxiliar |
| R24 | [RN]+[RV] | **CLIA** exige `FLAG_SEL_TRANSF_CLIA` | `TB_CNTR_BL` |

---

## Validação e execução da movimentação

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R25 | [RT]+[RN] | Pré-validação: `fnMovimentaCntr` deve retornar **OK** | `tbvw.Fields(0)` |
| R26 | [RN] | Persistência: `pcMovimentaCntr` com lacres, tara, usuário, motivo, câmera | `ADODB.Command` |
| R27 | [RN] | Divergência entre lacres digitados e manifesto pode gerar **avaria automática** (`local=12`, `tipo=32`, `complemento=8`) | Bloco `FlagDivLacre` |
| R28 | [RN]+[RV] | Se houver IMO e **`Verifica_Regras_Seg_Imo_Delta`** retornar texto, executar **segunda** `pcMovimentaCntr` **invertendo** origem/destino e alertar segregação | Final de `Atualiza_Posicao` |

---

## Avarias (aba dedicada)

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R29 | [RN]+[RV] | Inclusão exige **exatamente um** local e **≥1** tipo; sem seleção → erro | `cmdIncluirAvaria_Click` |
| R30 | [RT]+[RN] | Insert com `NOT EXISTS` para evitar duplicidade mesma combinação CNTR+LOCAL+TIPO | SQL dinâmico |
| R31 | [RN] | Exclusão só para linhas com autonum real; linhas “operador” (`autonum=0`) bloqueadas | `cmdExcluirAvaria_Click` + mensagem |
| R32 | [RN] | Finalização grava `TB_ETAPAS_AVARIA` com **cod_etapa=2** | `cmdFinalizarAvaria_Click` |
| R33 | [RN] | Se existe etapa **2, 3 ou 4** em `TB_ETAPAS_AVARIA`, considera **finalizado** — trava edições | `Verifica_Finalizado` |
| R34 | [RI]+[RT] | Permissões: função **998** habilita aba/botões; tipos 5/3/1 de permissão; **999** tipo 1 para finalizar | `Form_Load` |

---

## Grids de apoio (regras de filtro)

| ID | Classificação | Regra | Evidência |
|----|----------------|-------|-----------|
| R35 | [RT] | `Carrega_Grid2` distingue `Col_Patio=3` (usa `yard_eq`, substring diferente) dos demais | Múltiplos ramos |
| R36 | [RN]+[RT] | Para yards “valida=1”, a consulta inclui `VW_PROX_MOV_CNTR_ALL` + COL com LIKE progressivo; inclui lógica par/impar de **pilhas** vizinhas | trechos `Mod 2` |

---

## Hipóteses consolidadas

1. **[H]** Corpo de `pcMovimentaCntr` concentra transação e gravação em `TB_HIST_SHIFTING` / inventário — não lido aqui.
2. **[H]** `FrmVeiculoServ` pode ser obrigatório para completar dados de gate/saída — impacto direto em paridade funcional.
3. **[H]** `Trata_Erro_Atualizacao` pode estar órfão se `On Error` não apontar para o rótulo em todo o fluxo.

---

## Regras **não** encontradas (lacunas)

- Política completa de perfis para **movimentação** (além de avarias) — não há `Valida_Acesso_Botao` neste form; **hipótese**: controle só pelo menu.
- Regra exata de negócio por trás de cada código retornado por `fnMovimentaCntr` — depende do DDL da função no banco.
