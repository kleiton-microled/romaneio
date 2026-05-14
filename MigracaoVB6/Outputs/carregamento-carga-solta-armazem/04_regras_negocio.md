# Regras de Negocio - Carregamento Carga Solta (Armazem)

> Cada regra esta classificada como: `[NEG]` Regra de negocio, `[VAL]` Validacao, `[INT]` Regra de interface, `[TEC]` Regra tecnica, `[HIP]` Hipotese.

## Regras identificadas por acao

### Filtrar/Listar (Veiculo + Ordens + Carregamentos)
- `[NEG]` Carregamento e organizado por veiculo (cavalo + carreta) - apenas placas com OCs ativas em `VW_CAM_CARREGAMENTO` aparecem para selecao.
  - Evidencia: `Carrega_Veiculos` em `CarregaCS.frm`.
- `[NEG]` Apenas OCs vinculadas a placa selecionada e visiveis na lista `ORDENS`.
  - Evidencia: `WHERE A.PLACA_C='<placa8>'` em `Carrega_LvOrdem`.
- `[NEG]` A quantidade ja carregada de uma OC e calculada pela soma de `VOLUMES` de marcantes em `YARD='CAM'` cuja `PLACA_C` esta `NULL` (ainda nao bipado por outra placa) ou igual a placa atualmente selecionada.
  - Evidencia: subquery `B` no `Carrega_LvOrdem`.
- `[NEG]` Filtro por patio do operador: `Col_Patio` filtra `VW_CAM_CARREGAMENTO.PATIO`. `Col_Patio = 1` expande para `PATIO IN (1,7)` (regra historica - dois patios cooperando como um).
  - Evidencia: `Carrega_Veiculos`; `Col_Patio = 7 -> Col_Patio = 1` no login.
- `[NEG]` Carregamentos ja existentes (`lvCm`) sao apenas marcantes com `VOLUMES > 0` em `YARD='CAM'` para a mesma OC e mesma placa (ou ainda sem placa).
  - Evidencia: `Carrega_LvCarregamentos` SQL.
- `[INT]` Lista de "ORDENS" e ordenada por `LOTE, ITEM`.
- `[INT]` Lista de "CARREGAMENTOS" e ordenada por `M.AUTONUM` (cronologica).
- `[INT]` Combo de veiculo e ordenado por `DISPLAY` (placa cavalo).
- `[TEC]` `Carrega_LvOrdem` e disparado quando o veiculo muda no combo (`DataCombo1_Change`); `Carrega_LvCarregamentos` quando muda o veiculo, quando uma OC e clicada (`lvOC_Click`/`KeyUp`) ou apos carregar/cancelar.

### Buscar dados do marcante (bipagem)
- `[NEG]` Marcante so retorna dados se vinculado a uma `TB_CARGA_SOLTA` (importacao). Marcantes orfaos ou de exportacao nao popularem campos.
  - Evidencia: `INNER JOIN TB_MARCANTES M ON S.AUTONUM=M.AUTONUM_CARGA` em `Carrega_Dados`.
- `[NEG]` Os campos populados sao: `LOTE = S.BL`, `QTDE = M.VOLUMES`, `LOCAL = Y.YARD`, `ARMAZEM = A.DESCR`, `AUTONUMCS = S.AUTONUM`, `AUTONUM_CS_YARD = M.AUTONUM_CS_YARD`, `txtArmazem.Tag = S.ARMAZEM_IPA`.
- `[VAL]` Quando `Wflag_carregamento_marcante_bip <> 0` e o marcante foi digitado (Tag != "1" ou DateDiff > 1s), nao carrega dados; exibe modal "Campo nao permite digitacao manual" e limpa.
  - Evidencia: `txtMarcante_Change` e `txtMarcante_LostFocus`.
- `[NEG]` Quando `Wflag_carregamento_marcante_bip = 1`, apos validar marcante o campo e desabilitado ate `Limpa` (impede multiplos `F4` sem reinicializar contexto).
- `[INT]` Marcante deve ter exatamente 9 digitos para disparar `Carrega_Dados` automaticamente.
  - Evidencia: `If Len(Me.TXTMARCANTE) = 9 Then ...`.

### Carregar item (`F4`)
- `[VAL]` Item de OC selecionado e obrigatorio: senao "Ordem de carregamento nao selecionada".
- `[VAL]` Marcante obrigatorio: senao "Marcante nao informado/invalido".
- `[VAL]` AUTONUMCS do marcante deve coincidir com AUTONUMCS da OC selecionada: senao "Carga Divergente da Ordem de carregamento".
- `[VAL]` Marcante nao pode estar carregado: `LOCAL <> 'CAM'`. Se ja estiver, "Carga ja carregada".
- `[VAL]` Quantidade pedida nao pode exceder a quantidade original do marcante (`txtQtde > txtQtde.Tag`): "Quantidade indisponivel".
- `[NEG]` Funcao `FC_VALIDA_SAIDA_CARGA(BL, ORDEM, AUTONUMCS)` deve retornar `'OK'` para liberar; qualquer outro retorno e exibido como critica e a operacao e abortada.
  - Evidencia: SQL `SELECT DBO.FC_VALIDA_SAIDA_CARGA(...) AS RETORNO`.
- `[NEG]` O `UPDATE TB_MARCANTES SET ORDEM=..., PLACA_C='<placa>'` ocorre fora da transacao (antes do `BeginTrans`), associando logicamente o marcante ao veiculo e a OC mesmo antes de gravar a movimentacao yard.
  - Evidencia: linhas 816-818 de `cmdF4_Click` em `CarregaCS.frm`.
- `[NEG]` Carregamento total (`txtQtde == txtQtde.Tag`):
  - Zera o registro yard atual: `UPDATE TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=<id_yard_atual>`.
  - Insere registro com `YARD='CAM'`, `ORIGEM='I'`, `QUANTIDADE=<txtQtde>`, `MOTIVO=8`.
  - Atualiza marcante com novo `AUTONUM_CS_YARD`.
- `[NEG]` Carregamento parcial (`txtQtde < txtQtde.Tag`):
  - Decrementa do registro yard atual: `UPDATE TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE - <txtQtde>`.
  - Insere registro com `YARD='CAM'`, `QUANTIDADE=<txtQtde>`, `MOTIVO=8`.
  - Atualiza marcante: `AUTONUM_CS_YARD = novo`, `VOLUMES = <txtQtde>` (nota: o marcante "encolhe").
  - Atualiza `tb_temp_marcante.QTDE` para o mesmo valor onde `MARCANTE = <texto>`.
- `[NEG]` Sempre registra historico em `TB_HIST_SHIFTING_CS` com `MARCANTE`, `ARMAZEM=0`, `YARD='CAM'`, `DT_MOV=GETDATE()`, `USUARIO=Cod_Usuario`.
- `[INT]` Mensagem em verde no `txtCritica`: "ITEM CARREGADO" (total) ou "ITEM PARCIALMENTE CARREGADO" (parcial).
- `[NEG]` Apos sucesso: recarrega `lvCm`, recarrega `lvOC` e mantem a selecao na OC anterior (`OrdemAnt`); chama `Limpa` e devolve foco ao `txtMarcante`.

### Cancelar carregamento (`F5`)
- `[VAL]` Marcante obrigatorio: senao "Marcante nao informado/invalido".
- `[VAL]` Marcante deve estar em `'CAM'`: se nao estiver, "Carga ainda nao carregada".
- `[NEG]` Movimento e refletido como insert de novo registro com `YARD='CANCC'` (cancelamento) e simultaneo zeramento do registro `'CAM'` atual:
  - `UPDATE TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=<id_yard_atual>`.
  - `INSERT INTO TB_CARGA_SOLTA_YARD (..., YARD='CANCC', ORIGEM='I', QUANTIDADE=<txtQtde>, MOTIVO=8)`.
- `[NEG]` Marcante atualiza `AUTONUM_CS_YARD = novo` e `PLACA_C = NULL` (libera o marcante de qualquer placa).
- `[INT]` `Limpa` + foco em `txtMarcante`; recarrega `lvCm` e `lvOC`.
- `[NEG]` `F5` NAO insere em `TB_HIST_SHIFTING_CS` (intencional ou bug? - hipotese de bug; o cancelamento nao gera evento historico).
- `[HIP]` `'CANCC'` provavelmente significa "cancelado/cancelamento". Verificar dicionario de yards.

### Documentos (`F2`)
- `[VAL]` `txtLote > 0` obrigatorio: senao "Escolha uma carga".
- `[NEG]` Se a carga e LTL (`TB_SOLICITACAO_LTL.flag_ltl = 1` para o BL): abre `ConsultaDocColetor.exe` passando o BL.AUTONUM e `VLTL=true`.
- `[NEG]` Caso contrario (carga padrao): pega `MAX(autonum)` em `TB_AG_CS WHERE LOTE=<bl>` e abre o exe com esse id e `VLTL=false`.
- `[TEC]` Parametros do exe: `TrabDsn UsuBanco PasBanco Cod_Usuario <ID> TestaEntrada=true VLTL`.
  - Evidencia: `Shell(...)` em `cmdF2_Click`.

### Limpar (`F3`)
- `[INT]` Limpa `txtQtde`, `TXTLOTE`, `txtAutonum_Cs_Yard`, `txtAutonumCS`, `txtArmazem`, `txtArmazem.Tag`, `TXTLOCAL`, `txtCritica`, `TXTMARCANTE`. Reabilita `TXTMARCANTE`. Reseta `Tag` e os contadores anti-digitacao (`HorarioNumero`, `HorarioEnter`).

### Locais (`F1`)
- `[NEG]` Exige uma OC selecionada para abrir o painel.
- `[NEG]` Mostra inventario do armazem (`VW_INVENT_ARMAZEM`) com `qtde > 0` filtrado por `AUTONUMCS` da OC selecionada.
- `[INT]` Quando aberto, captura `ESC` para fechar (e somente fechar - o ESC fora deste estado fecha a tela).
- `[INT]` Lista ordenada por `DESCR_ARMAZEM, POSICAO, MARCANTE`.

### Sair / Navegar
- `[INT]` `ESC` fecha a tela quando o painel `frmLocais` esta oculto.
- `[INT]` `ESC` fecha apenas o painel `frmLocais` quando ele esta visivel.
- `[NEG]` Acesso a tela exige permissao via `Valida_Acesso_Botao` para `'COMMAND1(7)'` no sistema `'COLETOR_ARM'`.

## Validacoes
| Campo | Condicao | Mensagem |
|---|---|---|
| `Veiculo (DataCombo1)` | obrigatorio para listar OCs | nao ha mensagem - apenas a lista fica vazia |
| `OC (lvOC.SelectedItem)` | obrigatoria para `F4` e `F1` | "Ordem de carregamento nao selecionada" |
| `txtMarcante` | obrigatorio para `F4`/`F5`; tem 9 digitos | "Marcante nao informado/invalido" |
| `txtMarcante` (anti-digitacao) | quando `flag_bip <> 0`, exige bipagem | MsgBox "Campo nao permite digitacao manual" |
| `AUTONUMCS` da OC | igual ao `AUTONUMCS` do marcante | "Carga Divergente da Ordem de carregamento" |
| `LOCAL` (yard atual) | `<> 'CAM'` para `F4` | "Carga ja carregada" |
| `LOCAL` (yard atual) | `= 'CAM'` para `F5` | "Carga ainda nao carregada" |
| `txtQtde` | `<= txtQtde.Tag` | "Quantidade indisponivel" |
| `FC_VALIDA_SAIDA_CARGA` | retorno `= 'OK'` | retorno textual da function |
| `txtLote` (`F2`) | `> 0` | "Escolha uma carga" |
| `BL` da carga | encontrado em `TB_CARGA_SOLTA` | "Lote nao encontrado" |

## Regras de consistencia
- Quantidade:
  - `[NEG]` `M.VOLUMES` (do marcante) e a "porcao maxima" disponivel para carregamento daquela bipagem.
  - `[NEG]` Em carregamento parcial, o marcante "encolhe" (`VOLUMES = qtde_carregada`); o resto da quantidade fica como "registro yard nao zerado" no armazem.
  - `[HIP]` Por consequencia, marcantes parcialmente carregados podem precisar de uma nova bipagem para completar - hipotese a confirmar com usuario final.
- Integridade referencial:
  - `[TEC]` `TB_MARCANTES.AUTONUM_CS_YARD` aponta para `TB_CARGA_SOLTA_YARD.AUTONUM`. A cada operacao, um novo registro yard e criado (rastreabilidade) e o marcante e religado.
  - `[TEC]` `TB_MARCANTES.AUTONUM_CARGA` aponta para `TB_CARGA_SOLTA.AUTONUM` - chave de validacao de "Carga Divergente".
- Concorrencia:
  - `[TEC]` `Autonum_Cs_Yard = SELECT ISNULL(MAX(AUTONUM),0) FROM TB_CARGA_SOLTA_YARD WHERE AUTONUM_CS=<AUTONUMCS>` apos o INSERT - sob alta concorrencia em uma mesma OC, dois operadores podem trocar IDs. Requer revisao na migracao (usar `OUTPUT INSERTED.AUTONUM` ou `SCOPE_IDENTITY()`).
  - `[TEC]` Transacao envolve apenas yard + marcantes. O `UPDATE TB_MARCANTES SET ORDEM=..., PLACA_C=...` esta fora dela.

## Efeitos colaterais
- `[NEG]` Atualizacao indireta de `tb_temp_marcante.QTDE` quando ha carregamento parcial - dependencia da tela "Associacao de Marcantes" que tambem manipula essa tabela.
- `[NEG]` `TB_MARCANTES.PLACA_C` e gravada com 8 caracteres (Left do display) - "marca" o veiculo que carregou o marcante.
- `[NEG]` Cancelamento (`F5`) reseta `PLACA_C = NULL`, liberando o marcante para outras placas.
- `[TEC]` `TB_HIST_SHIFTING_CS` recebe somente o `F4`; `ARMAZEM=0` e `YARD='CAM'` (perde-se o armazem origem na escrita do historico - hipotese de simplificacao deliberada).
- `[TEC]` `Adodc1` carrega o combo via OLEDB (`ConnectionString` separado), enquanto o restante usa `db` (ADODB ODBC) - duplo canal.

## Hipoteses e duvidas
- `[HIP]` Funcao `FC_VALIDA_SAIDA_CARGA` engloba todas as regras de bloqueio de saida (free-time, GLME, SISCARGA, SEFAZ, BL, perfil, etc.). Deve ser preservada como-esta na migracao MVC. Acao: confirmar com DBA/CRC e mapear que tipos de criticas retornam para tratamento especifico no front (ex.: badges de bloqueio).
- `[HIP]` `MOTIVO=8` fixo em ambos `F4` e `F5` representa "carregamento por marcante" (ou similar). Acao: validar com tabela de motivos da `TB_CARGA_SOLTA_YARD`.
- `[HIP]` `'CAM'`/`'CANCC'` sao codigos de yard "logicos" (caminhao / cancelado). Acao: confirmar dicionario.
- `[HIP]` Quando a tela esta em modo bipagem (`flag_bip <> 0`), o leitor envia digitos + Enter. A regra de DateDiff > 1s identifica digitacao manual. Acao: testar com leitores reais para garantir que `keydown` capture todos os caracteres antes do `Enter`.
- `[HIP]` Carregamento parcial e raro, mas suportado. Acao: confirmar fluxo operacional.
- `[HIP]` `tb_temp_marcante` e a mesma tabela usada na "Associacao de Marcantes". O update parcial pode interferir naquela tela. Acao: revisar com time de operacao.
- `[HIP]` O update fora-de-transacao (`SET ORDEM, PLACA_C`) pode causar inconsistencia em caso de erro pos-update mas pre-commit. Acao: avaliar mover para dentro da transacao na migracao MVC.
- `[HIP]` Falta de historico em `F5` (cancelamento) parece bug do legado - sugere-se gravar evento `'CANCC'` em `TB_HIST_SHIFTING_CS` na MVC para rastreabilidade.
- `[HIP]` Filtro `PATIO IN (1,7)` para `Col_Patio=1` parece regra especifica de cliente. Acao: validar se ainda e necessaria.
- `[HIP]` `ConsultaDocColetor.exe` em ambiente web nao se aplica diretamente; alternativa: redirecionar para tela de documentos do proprio sistema (se houver) ou abrir nova aba apontando para um endpoint que entrega o PDF (precisa-se mapear como esse exe identifica/recupera os PDFs).
