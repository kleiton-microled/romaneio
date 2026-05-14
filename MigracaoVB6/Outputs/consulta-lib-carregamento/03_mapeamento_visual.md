# Mapeamento Visual - Consulta Lib. Carregamento

## Inventario de componentes
- Labels:
  - Protocolo, Placa, CNTR, BL, Periodo, Lote, N Documento, Transportadora, Motorista, Placas, Status Siscarga, Status Sefaz, Localizacao Carga, Tipo Veiculo, Prot. Simultaneos, alertas de bloqueio e GLME.
- Inputs:
  - `mskProtocolo`, `mskPlaca`, `CNTR`, `TxtFiltroBL`, `txtPeriodo`, `txtLote`, `txtBL`, `txtTransp`, `txtMot`, `txtPlacas`, `txtSiscarga`, `txtSefaz`, `txtCpfMot`, `txtLocal`.
- Combos:
  - `cbTipoVeiculo` (DataCombo carregado por SQL).
- Grid/ListView:
  - Nao ha grid tradicional; usar tabela web para melhorar leitura.
- Listas:
  - `listaCarga` (itens de carga vinculados).
  - `lstSimultaneos` (protocolos simultaneos).
- Botoes:
  - `Filtrar`, `Registrar`, `Limpar`, `Documentos`.
- Secoes/Paineis:
  - Bloco de filtros.
  - Bloco de dados de carga.
  - Bloco de status/liberacao.
  - Bloco de listas auxiliares.
  - Rodape com acoes.

## Equivalencia VB6 -> MVC
- Controle VB6: `MSMask MaskEdBox` (`mskProtocolo`, `mskPlaca`, `CNTR`)
  - Equivalente web: `input` com mascara JS (ex.: `imask` ou script proprio).
  - Observacao: manter padrao visual e semantico das mascaras para reduzir erro operacional.
- Controle VB6: `DataCombo cbTipoVeiculo`
  - Equivalente web: `select` com options vindas de endpoint.
  - Observacao: valor selecionado deve persistir ao carregar carga.
- Controle VB6: `ListBox listaCarga`
  - Equivalente web: tabela simples (`table`) ou `ul` rolavel.
  - Observacao: exibir origem (CS/CNTR) e descricao amigavel.
- Controle VB6: `ListBox lstSimultaneos`
  - Equivalente web: card/lista lateral.
  - Observacao: so exibir quando protocolo preenchido e houver dados.
- Controle VB6: Labels de bloqueio/GLME
  - Equivalente web: badges/alerts (`alert-warning`/`alert-danger`).
  - Observacao: exibicao condicional com destaque visual.

## Proposta de layout web
- Estrutura geral:
  - Linha 1: filtros (protocolo, placa, cntr, bl) + botao filtrar.
  - Linha 2: dados principais (periodo/lote/BL/transportadora/motorista/placas/local/tipo veiculo).
  - Linha 3: status SISCARGA/SEFAZ + alertas bloqueio/GLME.
  - Linha 4: listas de carga e simultaneos.
  - Linha 5: acoes (registrar, limpar, documentos, sair).
- Agrupamento de secoes:
  - "Consulta", "Dados da Carga", "Status de Liberacao", "Listas de Apoio", "Acoes".
- Hierarquia visual:
  - Filtro e status com maior destaque.
  - Campos informativos em estado read-only.

## Comportamentos visuais dinamicos
- Exibicao condicional:
  - botao registrar oculto/conforme flag de patio.
  - labels de bloqueio e GLME somente quando regra dispara.
  - lista de simultaneos so com protocolo valido.
- Estado de habilitado/desabilitado:
  - registrar desabilita quando falta dado minimo ou ha bloqueio impeditivo.
  - botao documentos habilita somente com carga validada.
- Selecao e foco:
  - foco inicial em placa (como no fluxo atual de limpeza/sucesso).
  - manter navegacao por teclado entre campos de filtro.

## UX e acessibilidade
- Melhorias sugeridas:
  - padronizar mensagens em area fixa de alerta (alem de modais).
  - destacar campos somente leitura com estilo consistente.
  - aplicar feedback de carregamento durante consultas.
- Riscos de usabilidade:
  - excesso de mensagens modais pode quebrar fluidez operacional.
  - mascaras estritas podem dificultar uso mobile (avaliar escopo desktop web).
- Atalhos e navegacao por teclado:
  - preservar `ESC` para sair/voltar.
  - mapear atalhos de acao (se definidos com usuario) sem conflitar com browser.
