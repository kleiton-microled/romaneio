# Regras de negocio, validacao, interface e tecnicas - Manutencao Marcantes

Legenda de classificacao:

- **RN**: Regra de negocio
- **RI**: Regra de interface
- **RT**: Regra tecnica
- **RV**: Regra de validacao
- **HP**: Hipotese (sem evidencia direta ou dependente de definicao externa)

## Catalogo de regras

| ID | Descricao | Classificacao | Evidencia |
|----|-----------|---------------|-----------|
| R01 | Somente lotes presentes em `SGIPA..VW_WMS_DIVERGENCIA_QTDE` podem ser filtrados; caso contrario, bloquear com mensagem. | RN + RV | `cmdFiltra_Click` — `EOF` apos `Open` |
| R02 | O filtro usa o valor numerico `Val(TXTLOTE)` como `LOTE` na view. | RT + RV | Concatenacao com `Val()` |
| R03 | A quantidade de referencia exibida (`txtQtde`) vem de `QUANTIDADE_REAL` da view de divergencia. | RN | Atribuicao no `cmdFiltra_Click` |
| R04 | O identificador tecnico da carga utilizado na listagem e o `AUTONUM_CS` retornado pela view de divergencia. | RN | `txtAutonumCS = tb1!AUTONUM_CS` |
| R05 | A lista so inclui marcantes com `DT_ASSOCIACAO IS NOT NULL` para a carga (`C.AUTONUM = AUTONUM_CS`). | RN | Clausula `WHERE` em `Carrega_Lv` |
| R06 | Posicao de armazem exibida somente quando existir registro em `OPERADOR..VW_INVENT_ARMAZEM` com `AUTONUMCS` corrente e `QTDE > 0`, casando `MARCANTE`. | RN | Subselect + `LEFT JOIN` |
| R07 | A soma das quantidades (`VOLUMES`) listadas e comparada a referencia (`txtQtde`); igual -> verde, diferente -> vermelho. | RI + RN (indicador) | Final de `Carrega_Lv` |
| R08 | A gravacao altera apenas `TB_MARCANTES.VOLUMES` do `AUTONUM` selecionado. | RN + RT | `UPDATE` no `cmdSalvar_Click` |
| R09 | O marcante exibido na lista e formatado com 9 digitos (zeros a esquerda) na coluna, mas o valor gravado usa `Val(txtMarcante)`. | RT + RI | `Right("000000000" & ...)` e `UPDATE` |
| R10 | Acesso ao menu depende de permissao `COMMAND1(10)` no sistema `COLETOR_ARM`. | RN + RT | `Valida_Acesso_Botao` em `Principal.Form_Activate` |
| R11 | No modo de menu REDEX, o botao de entrada fica invisivel; no IPA, visivel. | RI + RT | `Command1(10).Visible` em `Form_Activate` |
| R12 | Atalhos F4/F5/ESC disparam as mesmas acoes dos botoes. | RI | `Form_KeyDown` |
| R13 | Filtrar exige lote informado; salvar exige marcante e quantidade informados. | RV | Mensagens em `cmdFiltra_Click` / `cmdSalvar_Click` |
| R14 | Nao ha validacao no form para impedir quantidade negativa, zero, ou maior que saldo — **nao encontrado**. | HP | Ausencia de checagens adicionais |
| R15 | A view `VW_WMS_DIVERGENCIA_QTDE` pode representar divergencia WMS distinta de "BL autonum = campo LOTE" — significado exato do campo `LOTE` depende da view. | HP | View nao versionada no repositorio |

## Regras de persistencia e consistencia

| Topico | Detalhe | Classificacao |
|--------|---------|---------------|
| Transacao | Nao ha `BeginTrans`/`CommitTrans` no fluxo de salvar analisado. | RT |
| Auditoria | Nao ha insert em tabela de historico neste form. | RT (ausencia) |
| Concorrencia | Dois usuarios podem alterar o mesmo `AUTONUM` — sem lock otimista/pessimista evidenciado. | HP + Risco |

## Impacto na manutencao de marcantes

- Qualquer alteracao de `VOLUMES` e **imediata** no banco (`db.Execute`).
- **Risco de negocio**: ajuste incorreto pode distorcer inventario/carregamentos que dependem de `VOLUMES`.

## Separacao solicitada (resumo)

- **Negocio**: R01, R03–R08, R10, trecho de R07 (criterio de comparacao).
- **Interface**: R07 (cores), R09 (formatacao), R11, R12.
- **Tecnica**: R02, R09 (uso de `Val`), transacao ausente, SQL dinamico.
- **Validacao**: R01 (implicitamente), R13.
- **Hipotese**: R14, R15, impactos cruzados em romaneio nao mapeados neste `.frm`.

## Checklist interno (Agente 04)

- [x] Regras por acao (filtrar/salvar/sair) identificadas.
- [x] Validacoes com mensagens mapeadas.
- [x] Consistencia quantitativa (soma vs referencia) mapeada como indicador.
- [x] Transacao documentada como ausente.
- [x] Hipoteses separadas de fatos.
