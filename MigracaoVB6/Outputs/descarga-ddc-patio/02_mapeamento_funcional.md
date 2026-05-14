# Mapeamento funcional — Descarga DDC (Patio)

## Finalidade da tela (negocio)

Registrar e manter **itens de descarga DDC** associados a um **veiculo (placa)** em processo de **carregamento cross-docking**, amarrando:

- o **conteiner** (`tb_cntr_bl` via `cntr` do registro de saida),
- o **gate** de entrada aberto em patio (`tb_gate_new`),
- **lotes (BL)** ativos vinculados ao conteiner,
- **quantidade** e **embalagem** descarregadas,
- indicacao se a **entrega esta finalizada** (SIM/NAO),
- e atualizar o **status DDC** do conteiner (`flag_ddc` + `inicio_ddc`).

A lista inicial de veiculos vem de ordens de carregamento com **`flag_ddc=1`** e **`flag_saida=0`**.

## Pre-condicoes

- Operador autenticado no Coletor Patio (contexto implicito do exe; **Hipotese**: permissoes especificas para este botao podem existir em outros modulos — nao ha checagem neste `.frm`).
- Existir registro em `tb_registro_saida_cntr` elegivel ao filtro de `Carrega_Placa`.
- Para operar itens: existir **gate IN** sem **gate OUT** para a placa da ordem (`cbPlaca_LostFocus`).

## Fluxo principal (feliz)

1. Abrir a tela pelo menu **Descarga DDC** (`PrincipalS`, `Command1(15)`).
2. O sistema carrega o combo de **placas** (placa + identificacao de conteiner) filtrado por DDC.
3. Operador seleciona **placa**; ao sair do campo (`LostFocus`), o sistema resolve **gate** e carrega **lotes** ativos.
4. Operador seleciona **lote** e aciona **F4 - Carregar** (ou botao equivalente).
5. Sistema preenche dados internos (`TxtCntr`, `TxtOC`, etc.), habilita a area de itens e exibe o **grid** de descargas ja cadastradas para o par **cntr + gate**.
6. Operador inclui ou edita itens: informa **quantidade**, **embalagem**, **entrega finalizada (SIM/NAO)** e grava com **F5 - Salvar**.
7. Ao salvar, alem do item, o sistema pode **marcar inicio DDC** no conteiner e, se entrega finalizada (SIM), atualizar **`flag_ddc=2`**.

## Fluxos alternativos e erros

| Situacao | Comportamento no legado |
|----------|-------------------------|
| Carregar sem placa ou sem lote | `MsgBox` informando obrigatoriedade |
| Placa sem gate de entrada aberto | Mensagem "Nao consta entrada para o veiculo informado"; encerra fluxo daquele foco |
| Limpar (`F2`) | Reseta selecao e desabilita area de itens |
| Novo (`F3`) com area desabilitada | Mensagem "Nao Dispinivel" |
| Excluir item com saida em `tb_amr_gate` | Bloqueia exclusao |
| Atualizar quantidade menor que soma de saidas | Bloqueia com mensagem |
| Tentar mudar embalagem com saida | Mensagem de atencao; **Hipotese**: update pode seguir sem alterar embalagem dependendo do ramo SQL (codigo ambiguo) |

## Dados consultados

- Ordens / registros de saida e conteiner (`tb_registro_saida_cntr`, `tb_cntr_bl`, `tb_ordem_carregamento`).
- Gate atual (`tb_gate_new`).
- BL/lotes ativos (`tb_amr_cntr_bl`, `tb_bl`).
- Itens de descarga (`tb_descarga_armazem`) + descricao de embalagem (`dte_tb_embalagens`).
- Saidas associadas (`tb_amr_gate.quantidade_saida`).

## Dados cadastrados / alterados / removidos

| Entidade | Operacao |
|----------|----------|
| `tb_descarga_armazem` | INSERT (novo item), UPDATE (edicao), DELETE (exclusao) |
| `tb_cntr_bl.flag_ddc` (+ `inicio_ddc`) | UPDATE conforme regras pos-gravacao |

## Campos de entrada vs somente leitura

- **Entrada**: selecao de placa e lote (combos); quantidade; embalagem; SIM/NAO finalizacao; selecao de linha no grid para edicao.
- **Somente leitura / oculto**: `TxtGate`, `TxtCntr`, `TxtOC`, `TxtAutonum`, etc. — suportam logica e filtros; usuario nao edita diretamente na UI principal.

## Acoes disponiveis (botoes / atalhos)

| Acao | Atalho | Efeito funcional |
|------|--------|------------------|
| Carregar | F4 | Valida e carrega cabecalho operacional + grid |
| Limpar | F2 | Reseta tela |
| Novo | F3 | Prepara novo item |
| Salvar | F5 | Persiste item e atualiza DDC do conteiner |
| Excluir | F6 | Remove item se permitido |
| Sair | ESC | Fecha |

## Mensagens exibidas (texto aproximado — encoding corrompido no arquivo)

- "Placa nao Informada" / "Placa nao informada"
- "Lote nao Informado" / "Lote nao informada"
- "Nao consta entrada para o veiculo informado"
- "Informe finalizacao de entrega"
- "Quantidade Invalida"
- "Embalagem nao Informada"
- "Gate nao localizado"
- "Quantidade nao pode ser inferior a quantidade ja saida"
- "Atencao - A embalagem nao pode ser alterada pois ja existe saida para este item"
- "Confirma exclusao do item ?"
- "Ja consta saida para este item - Exclusao Cancelada"
- "Nao Dispinivel" (atalhos de item com area desabilitada)

## Consulta vs operacao

A tela **nao e somente consulta**: executa **INSERT/UPDATE/DELETE** em `tb_descarga_armazem` e **UPDATE** em `tb_cntr_bl`.

## Impacto em patio / DDC / conteiner / romaneio

- **Patio / gate**: amarra descarga ao **gate** atual (`tb_gate_new`) e a **placa** da ordem.
- **DDC / conteiner**: altera `flag_ddc` e pode setar `inicio_ddc` em `tb_cntr_bl`.
- **Romaneio / carregamento**: **Hipotese**: impacto indireto via tabelas de operacao; **nao** ha referencia a romaneio neste `.frm`.
