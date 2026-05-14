# Template 02 - Mapeamento Funcional

## Finalidade da tela
- Objetivo de negocio:
  - Associar marcantes (identificadores fisicos) a itens de carga solta ainda nao associados, com controle de quantidade.
- Usuario-alvo:
  - Operador de coletor com permissao do menu de marcantes (`COMMAND1(8)` no legado).

## Acoes do usuario
- Acao principal 1:
  - Filtrar cargas por Lote e/ou CNTR para selecionar um item de carga elegivel.
- Acao principal 2:
  - Incluir marcantes com quantidade em lista temporaria e salvar a associacao final.
- Acoes secundarias:
  - Remover marcante da lista temporaria.
  - Sair da tela.
  - Operar por atalhos de teclado (F1/F2/F4/F5/ESC).

## Fluxo funcional
1. Usuario informa filtro (Lote e/ou CNTR) e aciona Filtrar.
2. Sistema carrega grid de cargas sem marcante; se nao achar no fluxo normal, tenta fluxo exportacao e sinaliza `EXP`.
3. Usuario seleciona uma linha do grid; sistema carrega lista temporaria associada a aquela linha.
4. Usuario informa marcante + quantidade e adiciona na lista temporaria.
5. Usuario repete inclusao/remocao ate compor lista desejada.
6. Usuario salva; sistema valida consistencia de quantidades e grava associacao final em `TB_MARCANTES`.
7. Sistema exibe sucesso e limpa area de trabalho.

## Regras de uso percebidas
- Obrigatorio:
  - Informar marcante para adicionar.
  - Informar quantidade para adicionar.
  - Manter soma das quantidades da lista igual a quantidade do item selecionado para salvar.
- Opcional:
  - Informar lote e/ou CNTR para filtrar (sem filtro mostra base nao historica).
- Restricoes:
  - Marcante deve existir e nao pode estar previamente associado (`DT_ASSOCIACAO`).
  - Marcante de exportacao exige contexto EXP e compatibilidade de `AUTONUM_CEXP`.

## Mensagens funcionais
- Sucesso:
  - `MARCANTES ASSOCIADOS COM SUCESSO`
- Alerta:
  - `Informe o marcante`
  - `Informe a quantidade associada ao marcante`
  - `Marcante ja adicionado`
  - `Quantidades divergentes`
- Erro:
  - `Nr de marcante invalido`
  - `Nr de marcante ja utilizado`
  - `Nr de marcante de carga de exportacao`
  - `Marcante divergente ao lote exp.`
  - `Carga nao selecionada`

## Classificacao funcional das funcionalidades
- Funcionalidade principal:
  - associar marcantes a uma carga selecionada com controle de quantidade.
- Funcionalidade secundaria:
  - filtro por lote/cntr; staging de lista temporaria.
- Funcionalidade de apoio operacional:
  - atalhos de teclado para operacao rapida no coletor.

## Pontos de atencao
- Dependencias operacionais:
  - depende de dados de carga sem marcante nas views SGIPA.
  - depende da existencia/estado de marcante em `TB_MARCANTES`.
- Casos de borda:
  - nenhuma linha selecionada no grid durante adicionar/remover/salvar.
  - divergencia entre quantidade do item e soma da lista.
  - condicao de capacidade (`QtdeJaMarcada + qteM <= QtdeCs`) pode bloquear atualizacao sem mensagem especifica por item.
- Hipoteses:
  - comportamento de setas laterais na imagem nao esta explicito no codigo do form.
