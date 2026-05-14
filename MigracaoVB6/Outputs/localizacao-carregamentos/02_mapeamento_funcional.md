# Mapeamento Funcional - Localizacao Carregamentos

## Finalidade da tela
- Objetivo de negocio: localizar rapidamente os carregamentos em patio, mostrando status operacional de veiculo, carga, posicao e responsavel de liberacao.
- Usuario-alvo: operador de patio / equipe de carregamento.

## Acoes do usuario
- Acao principal 1:
  - Consultar lista de carregamentos ativos no patio (com filtro opcional por placa/carreta).
- Acao principal 2:
  - Selecionar uma linha para visualizar detalhes (cavalo, carreta, gate in, carga, quadra, rua, transportadora, responsavel).
- Acoes secundarias:
  - Atualizar resultados pelo botao `Filtrar`.
  - Fechar tela com `ESC`.
  - Na variante `S`: acionar `SAIDA` para conteiner com validacoes.

## Fluxo funcional
1. Sistema identifica patio do usuario (`Col_Patio`) e carrega descricao do patio.
2. Sistema consulta carregamentos recentes sem `gate_out`, aplicando filtro por patio e opcionalmente por placa/carreta.
3. Usuario seleciona linha na grid; tela exibe detalhe do registro selecionado.
4. (Somente `LocalizacaoS`) Se registro for conteiner (`FLAG_CNTR=1`), usuario pode acionar `SAIDA`.
5. Sistema valida condicoes de saida, solicita confirmacao e persiste movimentacao para `SAIDA`.

## Regras de uso percebidas
- Obrigatorio:
  - Patio de sessao deve existir para restringir consulta.
  - Somente registros com `dt_gate_out is null` entram no resultado.
- Opcional:
  - Filtro por placa/carreta.
- Restricoes:
  - Janela temporal de consulta e limitada (1 dia em `Localizacao.frm`; 3 dias em `LocalizacaoS.frm`).
  - Acao `SAIDA` nao aparece para carga solta (somente para conteiner).

## Mensagens funcionais
- Sucesso:
  - Nao ha mensagem de sucesso explicita para consulta; feedback por atualizacao visual da grid.
- Alerta:
  - "Conteiner ja posicionado em SAIDA".
  - "Existe um conteiner acima, impossivel movimentar (...)"
  - Confirmacao: "Confirma o carregamento da unidade".
- Erro:
  - Nao ha bloco de tratamento de erro estruturado; erros de banco tendem a propagar via runtime VB6.

## Pontos de atencao
- Dependencias operacionais:
  - Forte dependencia de estado global (patio, usuario, parametros de pilha/yard).
  - Dependencia de views e tabelas de integracao entre operador/sgipa.
- Casos de borda:
  - Consulta vazia deve manter tela consistente (contadores zerados e sem erro).
  - Filtro de placa mascara pode nao cobrir todos os padroes de placa novos.
  - Em `SAIDA`, regra de empilhamento depende de composicao do yard por tamanho (`Col_Tam_Pos_Yard`).
