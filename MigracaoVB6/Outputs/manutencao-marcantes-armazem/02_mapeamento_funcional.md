# Mapeamento funcional - Manutencao Marcantes (Coletor Armazem)

## Finalidade da tela (negocio)

Permitir **ajustar a quantidade (`VOLUMES`) de um marcante** associado a uma carga solta que consta como **divergente** no WMS, com apoio de:

- referencia de quantidade esperada/correta vinda da view de divergencia;
- lista dos marcantes da mesma carga (`AUTONUM_CS`) que estao **associados** (`DT_ASSOCIACAO IS NOT NULL`) e com posicao de **inventario de armazem** (`OPERADOR..VW_INVENT_ARMAZEM`, `QTDE > 0`);
- indicador visual comparando a **soma das quantidades listadas** com a **quantidade de referencia** do filtro.

**Classificacao**: manutencao **operacional** (ajuste de volume em `TB_MARCANTES`), nao cadastro mestre de armazem nem de usuario.

## Pre-condicoes

- Usuario autenticado no Coletor com permissao para o botao mapeado como `"COMMAND1(10)"` no modulo `COLETOR_ARM` (`Valida_Acesso_Botao` em `mdlColetor.bas`).
- Ambiente **IPA** no menu principal (`Sistema_Coletor` diferente de `REDEX`), pois o botao 10 fica **invisivel** no fluxo REDEX — **Regra de interface / controle de acesso** evidenciada em `Principal.Form_Activate`.

## Fluxo principal (feliz)

1. Usuario abre a tela pelo menu **Manutencao - Marcantes** (`Principal`, indice 10).
2. Informa o **LOTE** (campo `TXTLOTE`).
3. Aciona **Filtrar** (botao ou **F4**).
4. Sistema localiza registro em `SGIPA..VW_WMS_DIVERGENCIA_QTDE` para o `LOTE` informado.
5. Sistema exibe **Embalagem**, **Quantidade de referencia** (`QUANTIDADE_REAL`) e preenche internamente `AUTONUM_CS`.
6. Sistema carrega a **lista de marcantes** com armazem/posicao/item/quantidade.
7. Usuario **clica** na linha do marcante a corrigir (campo `TXTMARCANTE` preenchido).
8. Usuario informa a **nova quantidade** em `TXTQTDEN`.
9. Aciona **Salvar** (botao ou **F5**).
10. Sistema executa `UPDATE` em `TB_MARCANTES.VOLUMES`, limpa marcante/quantidade digitados, **recarrega** a lista e tenta manter a mesma linha selecionada.

## Fluxos alternativos e erro

| Situacao | Comportamento evidenciado |
|----------|---------------------------|
| Filtrar sem lote | Mensagem: "Informe o lote". |
| Lote sem registro na view de divergencia | Mensagem: "Lote nao encontrado como divergente" (texto no `.frm` com caracteres corrompidos). |
| Salvar sem marcante selecionado | Mensagem: "Selecione o marcante". |
| Salvar sem quantidade nova | Mensagem: "Informe a quantidade". |
| ESC ou botao Voltar | Encerra o form (`Unload Me`). |

## Dados consultados

- Registro de **divergencia** por lote (`VW_WMS_DIVERGENCIA_QTDE`).
- **Marcantes** da carga (`TB_MARCANTES` + `TB_CARGA_SOLTA`) com filtro de **associacao** (`DT_ASSOCIACAO IS NOT NULL`).
- **Inventario de armazem** (`OPERADOR..VW_INVENT_ARMAZEM`) apenas para enriquecer linhas com `DESCR_ARMAZEM` e `POSICAO` quando houver saldo (`QTDE > 0`) para o `AUTONUMCS` corrente.

## Dados cadastrados / alterados / removidos

| Operacao | Evidencia |
|----------|-----------|
| **Alteracao** | `UPDATE SGIPA..TB_MARCANTES SET VOLUMES = ... WHERE autonum = ...` |
| Inclusao / exclusao | **Nao ha** `INSERT` ou `DELETE` no form. |

## Campos de entrada (usuario)

| Campo | Obrigatoriedade funcional |
|-------|---------------------------|
| Lote (`TXTLOTE`) | Obrigatorio para filtrar. |
| Nova quantidade (`TXTQTDEN`) | Obrigatorio para salvar. |
| Selecao na lista (`lv`) | Obrigatoria para salvar (define marcante alvo). |

## Campos somente leitura / derivados

- `txtEmbalagem`, `txtQtde` (referencia do filtro), `TXTTOT` (soma da lista), colunas da lista exceto interacao por selecao.

## Listas / grids

- `lv`: lista marcantes com colunas **ARMAZEM**, **POSICAO**, **ITEM**, **QTD**, **MARCANTE** (formato 9 digitos no legado), **AUTONUMCS** (largura 0 — oculta na UI mas presente na lista).

## Botoes e acoes

| Acao | Disparo |
|------|---------|
| Filtrar | Clique em `[F4]-FILTRA` ou tecla F4 |
| Salvar | Clique em `[F5]-SALVA` ou tecla F5 |
| Voltar / Sair | `[ESC]-VOLTAR`, ESC no form, ou fechamento |

## Mensagens exibidas

Listadas no diagnostico tecnico (`01_diagnostico_vb6.md`).

## Validacoes

- Lote informado para filtro.
- Marcante selecionado (texto nao vazio) para salvar.
- Quantidade nova informada (texto nao vazio) para salvar.

**Ausencia evidenciada no codigo**: validacao de valor minimo/maximo, consistencia com inventario, bloqueio por soma total, confirmacao "tem certeza", audit trail de alteracao de volume.

## Impacto em armazem / marcantes / romaneio / carregamento

- **Marcantes**: altera diretamente `TB_MARCANTES.VOLUMES` — impacta quantidade do marcante em todo processo que le essa coluna (carregamento, inventario, WMS).
- **Armazem**: leitura de posicao via `VW_INVENT_ARMAZEM`; **nao** grava nesta view/tabela no form.
- **Romaneio / carregamento**: **Hipotese** — impacto indireto se romaneio ou liberacoes usarem `VOLUMES` do marcante; requer confirmacao com regras de negocio e outros sistemas.

## Checklist interno (Agente 02)

- [x] Finalidade descrita em linguagem de negocio.
- [x] Acoes principais e auxiliares listadas.
- [x] Fluxo feliz descrito.
- [x] Fluxos alternativos/erro descritos.
- [x] Campos obrigatorios/opcionais inferidos a partir de validacoes.
- [x] Mensagens funcionais relevantes listadas.
