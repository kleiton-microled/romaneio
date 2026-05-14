# Consolidado — Descarga DDC (Patio)

## Resumo executivo

No legado **Coletor Pátio** (`ColetorPatioSmall.vbp`), a funcionalidade comercialmente referida como **"Descarga DDC"** abre o form VB6 **`FrmDescargaArmazem`**, apesar do **caption** da janela ser **"Descarga Armazem"** e o **titulo visual superior** exibir **"CARREGAMENTO DDC"**. A tela gerencia **itens de descarga** na tabela **`tb_descarga_armazem`**, vinculados ao **conteiner** e ao **gate** corrente, e atualiza o **status DDC** em **`tb_cntr_bl`** (`flag_ddc` e `inicio_ddc`). O fluxo depende de **ordem de carregamento** com **`flag_ddc=1`** e **`flag_saida=0`**, e de **gate de entrada aberto** em **`tb_gate_new`**.

## Decisoes confirmadas (com base em codigo)

1. **Entrada de menu (Patio)**: `PrincipalS.frm` — `Command1(15)` com texto **"Descarga DDC"** chama `FrmDescargaArmazem.Show`.
2. **Projeto compilado Patio**: `ColetorPatioSmall.vbp` referencia `FrmDescargaArmazem.frm` e titulo **"Coletor Pátio"**.
3. **Nucleo da funcionalidade**: `FrmDescargaArmazem.frm` + `FrmDescargaArmazem.frx`.
4. **Persistencia de itens**: `tb_descarga_armazem` (INSERT/UPDATE/DELETE).
5. **Atualizacao de estado DDC do conteiner**: `tb_cntr_bl.flag_ddc` (+ `inicio_ddc` quando aplicavel).
6. **Bloqueio de exclusao / quantidade**: consultas a `tb_amr_gate.quantidade_saida`.

## Hipoteses em aberto

- Semantica completa dos valores de **`flag_ddc`** (0/1/2) e impactos em outros processos.
- Se ha **permissao** especifica para o botao (similar a outras telas do coletor).
- Como o binario real lida com **handlers orfaos** (`cbConteiner`, `DataGrid1`, `Command7`, `MskPlaca`) que nao aparecem na secao de desenho analisada.
- Equivalencia exata de SQL entre ambiente Oracle do legado e **SQL Server** do Romaneio.

## Itens bloqueantes antes de codificar

1. Confirmar **dialeto** e **nomes qualificados** das tabelas no ambiente alvo do Romaneio.
2. Resolver **titulacao UX** ("Descarga DDC" vs "Descarga Armazem" vs "CARREGAMENTO DDC") com negocio.
3. Clarificar **transacao** e tratamento de falha parcial ao salvar (item gravado mas `tb_cntr_bl` nao).
4. Revisar com DBA a logica de **UPDATE de embalagem** no legado (bloco aninhado com ramo aparentemente inalcancavel).

## Escopo da primeira implementacao sugerida

- Paridade de **consulta** (placas, lotes, itens, embalagens).
- **Carregar** + **Grid** + **Salvar** + **Excluir** com validacoes equivalentes.
- **Transacao** no salvar (melhoria em relacao ao legado).

## Criterios de aceite (migracao considerada adequada)

### Funcionais

- [ ] Lista de placas replica o filtro `flag_saida=0` e `flag_ddc=1` na ordem de carregamento.
- [ ] Ao selecionar placa, sistema resolve **gate** com entrada sem saida; caso contrario exibe mensagem equivalente ao legado.
- [ ] Lista de lotes retorna apenas BL com `flag_ativo=1`.
- [ ] **Carregar** exige placa e lote; preenche cabecalho operacional e habilita area de itens.
- [ ] Grid lista itens por **`cntr` + `gate`** correntes.
- [ ] **Salvar** exige quantidade > 0, embalagem, indicacao de finalizacao, gate e lote validos.
- [ ] **Salvar** aplica atualizacao de `tb_cntr_bl` conforme regras do legado (`flag_ddc` 0→1 com `inicio_ddc`; `2` quando finalizado SIM).
- [ ] **Excluir** bloqueia quando houver saida em `tb_amr_gate`.
- [ ] **Atualizar** bloqueia quantidade inferior a saida acumulada.

### Filtros e dados exibidos

- [ ] Colunas do grid equivalentes: quantidade, descricao embalagem, finalizado, identificadores internos conforme necessidade web (exibir ou ocultar IDs por decisao de UX).

### Acoes disponiveis

- [ ] Carregar, Limpar, Novo, Salvar, Excluir, Sair — com mensagens equivalentes.

### Validacoes

- [ ] Todas as validacoes listadas em `04_regras_negocio.md` com testes automatizados ou checklist manual assinado.

### Regras de descarga / DDC a preservar

- [ ] Vinculo item → `cntr`, `gate`, `lote`, `id_oc`.
- [ ] Progressao de `flag_ddc` e gravacao de `inicio_ddc` quando aplicavel.

### Comportamentos VB6 a manter vs adaptar

| Comportamento | Manter | Adaptar |
|---------------|--------|---------|
| Mensagens e bloqueios funcionais | X | |
| SQL concatenado | | X |
| Ausencia de transacao no salvar | | X |
| Titulos inconsistentes | | X (com aprovacao) |
| Dependencia de ADO/DataCombo | | X |

## Objetos de banco envolvidos (inventario)

### Tabelas

- `sgipa..tb_registro_saida_cntr` (ou equivalente qualificado via `Banco_Sgipa`)
- `sgipa..tb_cntr_bl`
- `sgipa..tb_ordem_carregamento`
- `sgipa..tb_descarga_armazem`
- `sgipa..tb_amr_cntr_bl`
- `sgipa..tb_bl`
- `sgipa..dte_tb_embalagens`
- `operador..tb_gate_new`
- `operador..tb_amr_gate`

### Views

- Nenhuma view explicita no codigo desta tela.

### Procedures / functions

- Nenhuma procedure explicita; uso de **sequence** `seq_tb_descarga_armazem.nextval` quando `Banco_Em_Uso = "ORACLE"`.

### Campos lidos / gravados (alto nivel)

- **Lidos**: placas, conteiner, ordem, flags de ordem, gate, lotes ativos, quantidades de saida, `flag_ddc` atual.
- **Gravados**: linhas em `tb_descarga_armazem`; `flag_ddc`, `inicio_ddc` em `tb_cntr_bl`.

### Joins e filtros obrigatorios

- Ver SQL citado em `01_diagnostico_vb6.md`.

### Transacoes / commits no legado

- Sem transacao explicita na rotina `Salvar`; multiplas chamadas `Executa`.

### Concorrencia / consistencia

- Dependencia forte do estado do **gate** e do **conteiner**; risco de uso concorrente sem lock logico.

### Impacto em status / localizacao / processo

- Atualiza **status DDC** do conteiner; demais impactos logisticos **nao evidenciados** neste form.

## Proposta inicial de migracao MVC (resumo)

Ver detalhes em `05_desenho_mvc.md`: `DescargaDdcPatioController` + `DescargaDdcPatioRepositorio` + `Index.cshtml` + `descarga-ddc-patio.js`, com endpoints JSON espelhando `Carrega_Placa`, resolucao de gate/lotes, `CARREGAR`, grid, `Salvar`, `Excluir`.

## Riscos e duvidas abertas

| Item | Tipo |
|------|------|
| Handlers orfaos / possivel divergencia fonte vs binario | RT |
| `Form_Unload` com `Unload Me` | RT |
| `txtMarca` sem controle | RT |
| Logica UPDATE embalagem / aninhamento de IF | RN/RT |
| Diferenca Oracle vs SQL Server | RT |
| Impacto indireto em romaneio | HP |

## Status de prontidao para implementacao

| Aspecto | Status |
|---------|--------|
| Localizacao e escopo no VB6 | **Pronto** |
| Levantamento SQL e regras principais | **Pronto** |
| Lacunas de codigo / divergencias | **Requer revisao** antes de codificar salvar |
| Alinhamento de banco com Romaneio | **Pendente** (dependencia de DBA) |
| UX de titulos | **Pendente** (dependencia de negocio) |

## Arquivos gerados nesta pasta

- `01_diagnostico_vb6.md`
- `02_mapeamento_funcional.md`
- `03_mapeamento_visual.md`
- `04_regras_negocio.md`
- `05_desenho_mvc.md`
- `06_plano_implementacao.md`
- `00_consolidado.md`
