# Template 00 - Consolidado da tela (Manutencao - Marcantes Armazem)

## Resumo executivo

- **Tela legada**: `AdminMarcante` (`ProjetoReferencia/Coletor/AdminMarcante.frm`), aberta pelo menu `Principal.Command1(10)` com caption **"Manutencao - Marcantes"** (acentos corrompidos no arquivo fonte).
- **Titulo interno do form**: faixa **"MANUTENCAO- MARCANTES"** (mesma ressalva de encoding). O sufixo **"Armazem"** nao aparece literalmente no menu/titulo analisado; a funcao e de **ajuste de `VOLUMES` de marcantes** em contexto de **divergencia WMS** e leitura de **inventario de armazem** (`OPERADOR..VW_INVENT_ARMAZEM`).
- **Escopo do estudo**: concluido (documentacao `01`–`06`); **sem implementacao MVC**.

## Decisoes confirmadas (baseadas em codigo)

1. O form **nao** insere/exclui marcantes; apenas **atualiza `TB_MARCANTES.VOLUMES`**.
2. O filtro inicial exige existencia em **`SGIPA..VW_WMS_DIVERGENCIA_QTDE`** pelo campo `LOTE` numerico da tela.
3. A lista de marcantes exige **`M.DT_ASSOCIACAO IS NOT NULL`** e amarra a carga por **`C.AUTONUM = AUTONUM_CS`** vindo da divergencia.
4. Atalhos: **F4** filtra, **F5** salva, **ESC** sai.
5. Permissao de menu no Coletor: funcao **`COMMAND1(10)`** em sistema **`COLETOR_ARM`**.

## Hipoteses em aberto

- **HP1**: Significado exato do campo `LOTE` na view `VW_WMS_DIVERGENCIA_QTDE` (pode nao ser o mesmo identificador exibido ao usuario em outras telas).
- **HP2**: Impacto em **romaneio/liberacao de carregamento** ao alterar `VOLUMES` sem outras validacoes.
- **HP3**: Comportamento real de `LV.SelectedItem` + `Val()` no ambiente VB6 runtime (codigo sugere fragilidade).

## Itens bloqueantes antes do go-live (migracao)

1. Inspecionar **`VW_WMS_DIVERGENCIA_QTDE`** no banco (campos, cardinalidade, significado de `LOTE`).
2. Alinhar **permissao web** equivalente a `COMMAND1(10)`.
3. Definir se a web deve **revalidar pertenca** marcante/carga no salvar (recomendado; **nao** existe no VB6).

## Criterios de aceite (migracao considerada concluida)

### Filtro

- Informar lote vazio -> mesma regra de bloqueio do legado ("Informe o lote").
- Lote sem registro na view de divergencia -> mensagem equivalente a "Lote nao encontrado como divergente".
- Lote valido -> exibir **embalagem** e **quantidade de referencia** (`QUANTIDADE_REAL`) e carregar lista com SQL equivalente ao legado.

### Grid

- Colunas visiveis: **ARMAZEM**, **POSICAO**, **ITEM**, **QTD**, **MARCANTE** (formato 9 digitos na exibicao).
- Dados de inventario de armazem apenas quando subconsulta retornar `QTDE > 0` para o par `(AUTONUMCS, MARCANTE)`.
- Incluir somente marcantes com **`DT_ASSOCIACAO IS NOT NULL`**.
- Calcular **total** da soma de `QTD` listada e aplicar cor: **verde** se igual referencia, **vermelho** se diferente, **preto** quando estado limpo (equivalente a `Limpa`).

### Selecao e edicao

- Selecionar linha define o **marcante alvo** (equivalente a `TXTMARCANTE` / `lv_Click`).
- Informar **nova quantidade** obrigatoria para salvar.

### Salvar

- Validar ausencia de marcante / quantidade com mensagens do legado.
- Persistir `UPDATE TB_MARCANTES SET VOLUMES = @novo WHERE AUTONUM = @marcante` com **parametros**.
- **Extensao recomendada na web**: impedir salvar se o marcante nao pertencer ao `AUTONUM_CS` corrente (fortalecimento de seguranca).
- Apos salvar: limpar campos de edicao e **recarregar** a lista; total e cores atualizados.

### Atalhos e navegacao

- F4/F5/ESC com mesma semantica funcional; documentar limites do navegador.

### Seguranca e auditoria (web)

- **Extensao recomendada**: mensagem de sucesso/erro; log de auditoria (nao evidenciado no VB6).

## Riscos e lacunas

| Item | Tipo |
|------|------|
| `UPDATE` direto sem transacao multi-tabela | Tecnico (baixo neste form) |
| SQL concatenado no VB6 | Tecnico / seguranca na migracao |
| Ausencia de validacao de dominio da nova quantidade | Negocio |
| Concorrencia entre usuarios | Negocio / operacao |
| Definicao da view divergencia fora do repo | Lacuna |
| Possivel SQL inconsistente se `Limpa` for chamado com `TXTLOTE` preenchido e `AUTONUM_CS` vazio | Lacuna tecnica legada |

## Escopo da implementacao (fase atual)

- **Incluido**: artefatos `01`–`06` e este consolidado.
- **Fora de escopo**: codigo C#/Razor/JS, alteracoes de banco, permissoes definitivas, testes automatizados.

## Proximas etapas

1. Validar **HP1** e view com DBA.
2. Implementar conforme `06_plano_implementacao.md`.
3. Homologar paridade com VB6 e ajustar textos de mensagem (encoding).

## Prontidao para implementacao

- **Nivel**: **Medio-Alto** — fluxo e SQL estao claros no `.frm`; lacunas concentram-se na **definicao da view** e em **permissoes/percursos web**.
- **Condicoes para inicio seguro**:
  - acesso de leitura a `VW_WMS_DIVERGENCIA_QTDE` e `OPERADOR..VW_INVENT_ARMAZEM` no mesmo ambiente do Romaneio;
  - decisao sobre reforco de validacao e auditoria na web.

## Indice dos artefatos

| Arquivo | Conteudo |
|---------|-----------|
| `01_diagnostico_vb6.md` | Arquivos, eventos, SQL, dependencias |
| `02_mapeamento_funcional.md` | Fluxos, dados, impacto |
| `03_mapeamento_visual.md` | Controles, equivalencia web, UX |
| `04_regras_negocio.md` | Catalogo classificado |
| `05_desenho_mvc.md` | Proposta MVC aderente ao Romaneio |
| `06_plano_implementacao.md` | Fases, testes, riscos |
