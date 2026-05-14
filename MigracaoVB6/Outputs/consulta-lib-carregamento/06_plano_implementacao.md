# Plano de Implementacao - Consulta Lib. Carregamento

## Etapas

### Etapa 1 - Base backend
- Escopo:
  - criar contratos (`ViewModel`, DTOs, interface de repositorio, modelos de resposta).
  - mapear e centralizar queries de consulta.
- Entregaveis:
  - classes de modelo e interface `IConsultaLiberacaoCarregamentoRepositorio`.
  - esqueleto do repositorio com metodos assinados.
- Criterio de pronto:
  - compilacao sem erros e contratos alinhados ao diagnostico.

### Etapa 2 - Fluxo de controller/endpoints
- Escopo:
  - implementar `Index`, `Filtrar`, `Registrar`, `CarregarSimultaneos`.
  - validar sessao e retorno JSON padrao do projeto.
- Entregaveis:
  - `ConsultaLiberacaoCarregamentoController`.
  - rotas e retorno de erros/sucesso padronizados.
- Criterio de pronto:
  - endpoints retornam payload esperado em testes manuais.

### Etapa 3 - View e scripts
- Escopo:
  - criar tela Razor com blocos funcionais mapeados.
  - implementar JS para filtros, render, mensagens e acoes.
- Entregaveis:
  - `Views/ConsultaLiberacaoCarregamento/Index.cshtml`.
  - `Content/js/consulta-liberacao-carregamento.js`.
- Criterio de pronto:
  - fluxo completo de consulta e preparacao de registro operando no browser.

### Etapa 4 - Persistencia e validacoes
- Escopo:
  - portar regras de negocio e validacoes criticas (patio, bloqueios, GLME, pagamento/free-time, janela/permissao, flags de patio).
  - encapsular registro em fluxo transacional.
- Entregaveis:
  - implementacao dos metodos de registro/liberacao.
  - logging de falhas e mensagens funcionais equivalentes.
- Criterio de pronto:
  - paridade comportamental com VB6 validada em cenarios essenciais.

### Etapa 5 - Refinamentos finais
- Escopo:
  - ajustar UX, mensagens, feedback de carregamento.
  - cobrir lacunas/hypoteses com validacao funcional.
- Entregaveis:
  - ajustes finais e checklist de homologacao.
- Criterio de pronto:
  - aceite funcional e tecnico documentado.

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
  - consulta por protocolo valido retornando dados completos.
  - consulta por placa + CNTR/BL retornando dados e listas.
  - registro concluido com retorno de sucesso.
- Cenarios de erro:
  - filtros insuficientes.
  - patio divergente.
  - bloqueio BL/CNTR.
  - pendencia GLME.
  - falha de integracao SISCARGA/SEFAZ.
- Cenarios de borda:
  - protocolo inexistente.
  - carga sem ordem de carregamento associada.
  - carga com LTL.
  - tentativa fora da janela de agendamento.

## Riscos e mitigacoes
- Risco:
  - regras centrais em funcoes VB6 externas pouco visiveis no form.
- Mitigacao:
  - validar/portar funcoes compartilhadas antes de liberar etapa 4.

- Risco:
  - divergencia de comportamento por excesso de `MsgBox` no legado.
- Mitigacao:
  - mapear cada mensagem para estado de negocio no frontend.

- Risco:
  - dependencia de sistema externo de documentos sem equivalente web imediato.
- Mitigacao:
  - tratar como funcionalidade opcional na primeira entrega e alinhar com operacao.
