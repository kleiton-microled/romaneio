# Template 06 - Plano de Implementacao

## Etapas

### Etapa 1 - Base backend
- Escopo:
  - criar `IAssociacaoMarcantesRepositorio` + `AssociacaoMarcantesRepositorio`.
  - criar DTOs de filtro, linha e staging.
  - implementar consultas de filtro (normal/CEXP) e conteiner por lote.
- Entregaveis:
  - repositorio com consultas parametrizadas.
  - testes de query basica (manual ou automatizado de integracao).
- Criterio de pronto:
  - endpoints conseguem devolver dados de grid e modo EXP com IDs tecnicos.

### Etapa 2 - Fluxo de controller/endpoints
- Escopo:
  - criar `AssociacaoMarcantesController` com `Index`, `Filtrar`, `CarregarTemporarios`, `AdicionarTemp`, `RemoverTemp`, `Salvar`.
  - padronizar respostas de sucesso/erro.
- Entregaveis:
  - contrato de payload e respostas JSON.
  - validacoes backend das regras centrais.
- Criterio de pronto:
  - fluxo backend ponta a ponta operando sem view final.

### Etapa 3 - View e scripts
- Escopo:
  - construir `Views/AssociacaoMarcantes/Index.cshtml`.
  - implementar JS de grid/staging/atalhos/feedback.
- Entregaveis:
  - tela navegavel com fluxo completo sem refinamento visual final.
- Criterio de pronto:
  - usuario filtra, seleciona, inclui/remove staging e salva pela UI web.

### Etapa 4 - Persistencia e validacoes
- Escopo:
  - consolidar regras de quantidade e contexto EXP.
  - transacao no salvar para reduzir parcialidade.
  - revisar mensagens de erro/sucesso.
- Entregaveis:
  - persistencia final validada com base de homologacao.
- Criterio de pronto:
  - dados gravados equivalentes ao legado nos cenarios principais.

### Etapa 5 - Refinamentos finais
- Escopo:
  - ajustes de UX, acessibilidade e ergonomia de atalhos.
  - revisao de seguranca SQL/logs.
  - documentacao final da feature.
- Entregaveis:
  - tela pronta para homologacao funcional com operacao.
- Criterio de pronto:
  - criterios de aceite atendidos e pendencias criticas zeradas.

## Checklist tecnico
- [ ] Backend
  - [ ] Repositorio com SQL parametrizado
  - [ ] Validacoes de negocio no servidor
  - [ ] Transacao no salvar
- [ ] Frontend
  - [ ] Grid selecionavel
  - [ ] Lista temporaria de marcantes
  - [ ] Atalhos F1/F2/F4/F5/ESC com fallback de botoes
- [ ] Banco
  - [ ] Confirmacao de esquema/indices de `TB_TEMP_*`
  - [ ] Confirmacao de estrategia de limpeza da temp
- [ ] Rotas
  - [ ] GET Index
  - [ ] POST de filtro/staging/salvar
- [ ] Validacoes
  - [ ] Marcante obrigatorio
  - [ ] Qtde obrigatoria
  - [ ] Regras EXP
  - [ ] Quantidades divergentes
- [ ] Tratamento de erros
  - [ ] Mensagens equivalentes ao legado
  - [ ] Sem erros silenciosos
- [ ] Testes
  - [ ] Fluxo feliz normal
  - [ ] Fluxo feliz EXP
  - [ ] Cenarios de erro principais

## Plano de testes
- Cenarios felizes:
  - filtrar -> selecionar -> adicionar marcantes -> salvar com sucesso (normal).
  - filtrar sem resultado normal -> fallback CEXP -> salvar com sucesso.
- Cenarios de erro:
  - marcante invalido, ja utilizado, divergente no EXP.
  - soma de quantidade divergente.
  - tentativa de salvar sem selecao valida.
- Cenarios de borda:
  - concorrencia de alteracao da quantidade associada por outro processo.
  - remocao com item de staging nao selecionado.
  - refresh da pagina durante operacao de staging.

## Riscos e mitigacoes
- Risco:
  - comportamento legado com `On Error Resume Next` escondia falhas.
- Mitigacao:
  - no MVC, registrar erro e retornar mensagem clara.

- Risco:
  - divergencia de comportamento de atalhos no navegador.
- Mitigacao:
  - mapear atalhos sob escopo da pagina e manter botoes equivalentes.

- Risco:
  - persistencia parcial no salvar por loop sem transacao.
- Mitigacao:
  - encapsular salvar em transacao e abortar em inconsistencias.

- Risco:
  - regra de limpeza de temp desconhecida.
- Mitigacao:
  - validar com DBA/operacao antes de fechar implementacao.

## Ordem sugerida de desenvolvimento
1. Backend de leitura/escrita e contratos.
2. Controller com validacoes e respostas.
3. View e scripts.
4. Regras finas e transacao.
5. Homologacao e ajustes.
