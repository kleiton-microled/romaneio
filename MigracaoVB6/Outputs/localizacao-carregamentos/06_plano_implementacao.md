# Plano de Implementacao - Localizacao Carregamentos

## Etapas

### Etapa 1 - Base backend
- Escopo:
  - Criar contratos (`interface`, DTOs, ViewModel) e repositorio com consulta principal.
- Entregaveis:
  - `ILocalizacaoCarregamentosRepositorio`
  - `LocalizacaoCarregamentosRepositorio` (metodo `Filtrar`).
- Criterio de pronto:
  - Consulta retorna dataset equivalente ao VB6 para patio/filtro.

### Etapa 2 - Fluxo de controller/endpoints
- Escopo:
  - Criar controller, endpoints `Index` e `Filtrar`, validacao de sessao.
- Entregaveis:
  - `LocalizacaoCarregamentosController`.
- Criterio de pronto:
  - Endpoint responde JSON padrao e bloqueia sem sessao.

### Etapa 3 - View e scripts
- Escopo:
  - Montar tela Razor com filtro, grid, painel de detalhe e totalizadores.
- Entregaveis:
  - `Views/LocalizacaoCarregamentos/Index.cshtml`
  - `Content/js/localizacao-carregamentos.js`.
- Criterio de pronto:
  - Fluxo de consulta e selecao de linha funcional no browser.

### Etapa 4 - Persistencia e validacoes
- Escopo:
  - Implementar opcao `SAIDA` (se confirmada para escopo) com validacoes e transacao.
- Entregaveis:
  - endpoint `MoverSaida`
  - validacoes de SAIDA e empilhamento.
- Criterio de pronto:
  - update/historico atomicos e mensagens equivalentes ao legado.

### Etapa 5 - Refinamentos finais
- Escopo:
  - Ajustes de usabilidade, performance, homologacao com operacao.
- Entregaveis:
  - melhoria de mensagens, mascaras, atalhos e logs.
- Criterio de pronto:
  - aceite funcional e tecnico concluido.

## Checklist tecnico
- [ ] Backend
- [ ] Frontend
- [ ] Banco
- [ ] Rotas
- [ ] Validacoes
- [ ] Tratamento de erros
- [ ] Testes

## Plano de testes
- Cenarios felizes:
  - Consultar sem filtro e com filtro por placa/carreta.
  - Selecionar linha e validar painel de detalhe.
  - (Se habilitado) executar `SAIDA` com sucesso.
- Cenarios de erro:
  - Sessao expirada.
  - Filtro invalido/sem retorno.
  - Tentativa de `SAIDA` em conteiner ja em SAIDA.
  - Tentativa de `SAIDA` com conteiner acima.
- Cenarios de borda:
  - Patio com regra 1/7.
  - Janela temporal no limite.
  - Divergencia entre variantes de consulta.

## Riscos e mitigacoes
- Risco:
  - Ambiguidade entre `Localizacao.frm` e `LocalizacaoS.frm` (query e regras diferentes).
- Mitigacao:
  - Congelar baseline com negocio/DBA antes de codar.

- Risco:
  - Inconsistencia no fluxo de `SAIDA` se nao houver transacao.
- Mitigacao:
  - Implementar transacao unica para update + historico.

- Risco:
  - Diferencas de UX entre grid VB6 e web.
- Mitigacao:
  - Validar prototipo com operador antes de homologacao final.
