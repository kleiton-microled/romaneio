# Plano de implementacao - Manutencao Marcantes Armazem

## Premissas

- Estudo concluido nos artefatos `01` a `05` e consolidado no `00_consolidado.md`.
- Nenhuma implementacao iniciada nesta fase.
- Banco `SGIPA` / `OPERADOR` acessivel pelo mesmo `Config.StringConexao()` ja usado pelo Romaneio.

## Fases sugeridas

### Etapa 1 — Contratos e acesso a dados

- Criar `IManutencaoMarcantesArmazemRepositorio` com metodos:
  - `ObterDivergenciaPorLote(int lote)` -> DTO ou nulo.
  - `ListarMarcantesPorAutonumCs(int autonumCs)` -> lista de linhas.
  - `AtualizarVolumesMarcante(int autonumMarcante, int novosVolumes, int autonumCs)` -> resultado (sucesso/erro).
- Implementar `ManutencaoMarcantesArmazemRepositorio` com **SQL parametrizado** (Dapper).
- Incluir **validacao de pertenca** do marcante ao `AUTONUM_CS` na mesma consulta do `UPDATE` (por exemplo, `UPDATE ... FROM TB_MARCANTES M JOIN TB_CARGA_SOLTA C ON ... WHERE M.AUTONUM = @Autonum AND C.AUTONUM = @AutonumCs`).

**Pronto desta etapa**: metodos compilam; consultas executam em ambiente de dev; testes manuais via script SQL ou teste de integracao basico (opcional).

### Etapa 2 — Controller e endpoints

- Criar `ManutencaoMarcantesArmazemController` com `Index`, `Filtrar`, `Salvar`.
- Padronizar checagem de sessao e retorno JSON como telas similares.
- Mapear erros de banco para mensagens amigaveis (sem vazar stack trace).

**Pronto**: chamadas Postman/browser retornam JSON coerente.

### Etapa 3 — View Razor e front-end

- Criar `Views/ManutencaoMarcantesArmazem/Index.cshtml` com layout alinhado a `AssociacaoMarcantes/Index.cshtml` (cards, botoes, tabela).
- Criar `Content/js/manutencao-marcantes-armazem.js` para:
  - atalhos F4/F5/ESC;
  - chamadas Ajax;
  - atualizacao da tabela e cores do total;
  - selecao de linha.

**Pronto**: fluxo ponta a ponta em ambiente local.

### Etapa 4 — Integracao no menu Romaneio

- Adicionar entrada em `Views/Home/Home.cshtml` (ou mecanismo de menu existente) apontando para `ManutencaoMarcantesArmazem/Index`.
- Definir permissao / perfil se o Romaneio tiver controle fino (espelhar `COMMAND1(10)` — **Hipotese**: pode exigir nova funcao no sistema de permissoes web; validar com negocio).

**Pronto**: usuario autenticado acessa pelo menu.

### Etapa 5 — Homologacao funcional

- Casos de teste derivados dos criterios de aceite do consolidado.
- Comparacao lado a lado com VB6 (mesmo lote, mesma lista, mesmo update) em ambiente controlado.

## Dependencias entre etapas

```text
Etapa 1 -> Etapa 2 -> Etapa 3 -> Etapa 4 -> Etapa 5
```

## Checklist tecnico (execucao futura)

- [ ] Interface + repositorio registrados no Unity (`Global.asax.cs`).
- [ ] Arquivos `.cs` incluidos no `Romaneio.csproj`.
- [ ] SQL sem concatenacao de entrada do usuario.
- [ ] Validacao de pertenca marcante/carga no salvar.
- [ ] Tratamento de excecao SQL com mensagem operacional.
- [ ] View responsiva minima (uso em coletor / tablet).
- [ ] JS: prevenir comportamento padrao problematico de F5.
- [ ] Textos de mensagem alinhados ao legado (com correcao de encoding).

## Plano de testes (alto nivel)

| Caso | Esperado |
|------|----------|
| Filtrar lote valido divergente | Retorna embalagem, referencia, lista e cor correta do total. |
| Filtrar lote inexistente | Mensagem equivalente a "nao encontrado como divergente". |
| Salvar sem selecao / sem quantidade | Mensagens de validacao. |
| Salvar marcante que nao pertence ao lote filtrado (payload adulterado) | **Bloqueio** com erro (comportamento web melhorado vs VB6). |
| Salvar quantidade valida | `VOLUMES` atualizado; lista e total recalculados. |

## Riscos por etapa

| Etapa | Risco | Mitigacao |
|-------|-------|-----------|
| 1 | Definicao da view divergencia desconhecida | Validar com DBA; inspecionar view em homologacao. |
| 2 | Exposicao de endpoint sem autorizacao | Reusar padrao de sessao + futura permissao fina. |
| 3 | Atalhos conflitantes no browser | Tooltips + botoes visiveis. |
| 4 | Permissao nao mapeada no web | Alinhar com TI antes do go-live. |

## Checklist interno (Agente 06)

- [x] Etapas pequenas definidas.
- [x] Dependencias explicitadas.
- [x] Criterios de pronto por etapa.
- [x] Plano de testes definido.
- [x] Riscos por etapa mapeados.
