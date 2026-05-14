# Mapeamento Visual - Localizacao Carregamentos

## Inventario de componentes
- Labels:
  - Titulo da tela com patio.
  - Rotulos de detalhe: cavalo, carreta, lib. ent., gate in, carga, quadra, rua, transp., responsavel.
  - Rotulos de totalizadores: CNTR, CS, TOT.
- Inputs:
  - Campos de detalhe somente leitura (`txtCavalo`, `txtCarreta`, `txtLibera`, etc.).
  - `TxtPatio` somente leitura.
- Combos:
  - Nao ha combo no form.
- Grid/ListView:
  - `MSHFlexGrid Grid1` com colunas de dados e flags internas.
- Listas:
  - Nao ha lista auxiliar dedicada.
- Botoes:
  - `Filtrar`.
  - `SAIDA` (apenas variante `S` e condicional por tipo de linha).
- Secoes/Paineis:
  - `Frame1` principal.
  - `Frame2` de filtro por placa.

## Equivalencia VB6 -> MVC
- Controle VB6: `MSMask mskPlaca`
  - Equivalente web: `input` com mascara JS (placa antiga/mercosul) + normalizacao no backend.
  - Observacao: preservar filtro em placa e carreta.
- Controle VB6: `MSHFlexGrid Grid1`
  - Equivalente web: tabela HTML responsiva com selecao de linha.
  - Observacao: manter colunas de negocio e esconder colunas tecnicas (flags/autonum) quando necessario.
- Controle VB6: `txt*` bloqueados
  - Equivalente web: campos readonly ou painel de detalhe.
  - Observacao: atualizacao por evento de selecao da linha.
- Controle VB6: `CmdSaida`
  - Equivalente web: botao de acao contextual habilitado apenas para conteiner.
  - Observacao: pedir confirmacao modal antes da persistencia.

## Proposta de layout web
- Estrutura geral:
  - Cabecalho com titulo + patio.
  - Bloco superior com resumo do item selecionado.
  - Bloco de filtro (placa + botao filtrar).
  - Grid principal em largura total.
  - Rodape com totalizadores.
- Agrupamento de secoes:
  - "Resumo selecionado", "Filtro", "Resultados", "Acoes".
- Hierarquia visual:
  - Priorizar grid e totalizadores, mantendo detalhe em leitura rapida.

## Comportamentos visuais dinamicos
- Exibicao condicional:
  - Botao `SAIDA` aparece somente para linhas de conteiner.
- Estado de habilitado/desabilitado:
  - Botao `SAIDA` desabilitado sem linha valida.
- Selecao e foco:
  - Clique em linha atualiza painel de detalhe e mantem destaque visual.

## UX e acessibilidade
- Melhorias sugeridas:
  - Mensagens de erro/sucesso padronizadas (alerta inline e modal de confirmacao).
  - Indicador de carregamento durante consulta.
  - Coluna de tipo (CNTR/CS) explicitada no grid para reduzir ambiguidade.
- Riscos de usabilidade:
  - Em VB6, parte da informacao fica "escondida" por largura de coluna; na web precisa ser explicita.
  - Mascara unica de placa pode rejeitar formatos atuais.
- Atalhos e navegacao por teclado:
  - Preservar `ESC` para sair.
  - Opcional: `Enter` no filtro para executar `Filtrar`.
