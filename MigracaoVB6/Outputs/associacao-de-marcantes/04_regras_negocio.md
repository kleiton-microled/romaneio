# Template 04 - Regras de Negocio

## Regras identificadas por acao

### Filtrar
- Regra: sem filtro (lote/cntr vazios), listar apenas `FLAG_HISTORICO=0` e `QUANTIDADE>0`.
- Evidencia: metodo `Carrega_Lv`.
- Tipo: Negocio + Tecnica.

- Regra: se a view principal nao retornar dados, tentar view de exportacao e ativar modo `EXP`.
- Evidencia: `If tb1.EOF Then ... VW_WMS_SEM_MARCANTE_CEXP ... LBLEXPORTACAO.Visible = True`.
- Tipo: Negocio.

### Incluir
- Regra: marcante e quantidade sao obrigatorios para inclusao em lista temporaria.
- Evidencia: validacoes iniciais de `cmdAdiciona_Click`.
- Tipo: Negocio.

- Regra: nao permitir marcante duplicado na mesma lista temporaria em tela.
- Evidencia: loop em `lstMarcantes` comparando os 9 primeiros caracteres.
- Tipo: Negocio.

- Regra: marcante precisa existir em `TB_MARCANTES` e nao pode estar associado (`DT_ASSOCIACAO` vazio).
- Evidencia: query em `TB_MARCANTES` + validacao `DT_ASSOCIACAO`.
- Tipo: Negocio.

- Regra: marcante de exportacao (`AUTONUM_CEXP > 0`) exige modo EXP e mesmo `AUTONUM_CEXP` da carga selecionada.
- Evidencia: validacoes no bloco `If Nnull(tb1!AUTONUM_CEXP, 0) > 0 Then`.
- Tipo: Negocio.

### Remover
- Regra: remover item da tabela temporaria com base em `AUTONUM_CS/AUTONUM_CEXP` da linha selecionada e `MARCANTE` selecionado na lista.
- Evidencia: `cmdRemove_Click`.
- Tipo: Tecnica (staging) com impacto funcional.

### Salvar/Persistir
- Regra: exige ao menos uma linha no grid para salvar.
- Evidencia: `If Me.lv.ListItems.Count = 0 Then ...`.
- Tipo: Interface/Tecnica.

- Regra: soma das quantidades da lista temporaria deve ser igual a quantidade do item selecionado.
- Evidencia: `qteM` (soma da lista) comparada com `qteItem`.
- Tipo: Negocio.

- Regra: so atualiza marcante se `QtdeJaMarcada + qteM <= QtdeCs`.
- Evidencia: condicional antes do `UPDATE`.
- Tipo: Negocio.

- Regra: grava campos adicionais no fluxo normal (`BL`, `ITEM`, `CNTR`); no fluxo EXP grava somente associacao CEXP e volumes.
- Evidencia: dois blocos distintos de `UPDATE`.
- Tipo: Negocio.

### Sair/Navegar
- Regra: ESC fecha o form.
- Evidencia: `Form_KeyDown` + `cmdEscape_Click`.
- Tipo: Interface.

## Validacoes
- Campo: Marcante
- Condicao: vazio/zero
- Mensagem: `Informe o marcante`
- Tipo: Negocio.

- Campo: Qtde
- Condicao: vazio/zero
- Mensagem: `Informe a quantidade associada ao marcante`
- Tipo: Negocio.

- Campo: Marcante
- Condicao: inexistente em `TB_MARCANTES`
- Mensagem: `Nr de marcante invalido`
- Tipo: Negocio.

- Campo: Marcante
- Condicao: `DT_ASSOCIACAO` preenchido
- Mensagem: `Nr de marcante ja utilizado`
- Tipo: Negocio.

- Campo: Contexto EXP
- Condicao: marcante de exportacao fora do modo EXP
- Mensagem: `Nr de marcante de carga de exportacao`
- Tipo: Negocio.

- Campo: Contexto EXP
- Condicao: divergencia de `AUTONUM_CEXP`
- Mensagem: `Marcante divergente ao lote exp.`
- Tipo: Negocio.

- Campo: Salvar
- Condicao: soma de quantidades divergente
- Mensagem: `Quantidades divergentes`
- Tipo: Negocio.

## Regras de consistencia
- Quantidade:
  - `sum(qtde_lista_temp) == quantidade_item_selecionado`
  - `qtde_ja_marcada + qtde_marcante_atual <= qtde_cs_disponivel`
- Integridade referencial:
  - associacao final aponta para `AUTONUM_CARGA` ou `AUTONUM_CEXP` conforme contexto.
- Concorrencia:
  - nao ha bloqueio transacional explicito no form; leituras e updates sao separados.

## Efeitos colaterais
- Atualizacoes indiretas:
  - no fluxo normal, update em `TB_MARCANTES` tambem grava `BL`, `ITEM`, `CNTR`.
- Reuso de tabela temporaria:
  - lista em tela depende de `TB_TEMP_MARCANTE` / `TB_TEMP_MARCANTE_CEXP`.
- Dependencias externas:
  - depende de status de carga e marcante em tabelas SGIPA em tempo real.

## Classificacao consolidada das regras por tipo
- Regras de negocio:
  - obrigatoriedade de campos, validade de marcante, compatibilidade EXP, consistencia de quantidades, limite de capacidade.
- Regras de interface:
  - atalhos de teclado, foco pos-acao, habilitacao de botoes conforme resultado.
- Regras tecnicas:
  - uso de tabelas temporarias, SQL dinamico, fallback de view, colunas ocultas no `ListView`.
- Hipoteses:
  - limpeza da temp apos salvar pode ocorrer fora da tela.
  - setas de navegacao da imagem nao estao codificadas no `.frm`.

## Hipoteses e duvidas
- Hipotese: existem processos externos que limpam `TB_TEMP_*` apos consolidacao.
- Evidencia parcial: o form nao executa `DELETE` geral de temp no salvar.
- Acao de validacao: confirmar trigger/job/procedure no banco ou em outros forms/modulos.

- Hipotese: a UI com setas laterais pertence a versao/asset nao presente no fonte textual.
- Evidencia parcial: controle nao encontrado no `Marcantes.frm`.
- Acao de validacao: inspecionar executavel legado e/ou outra branch/fonte historica.
