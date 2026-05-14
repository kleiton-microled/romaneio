# Mapeamento Funcional - Carregamento Carga Solta (Armazem)

## Finalidade da tela
- Objetivo de negocio: registrar fisicamente o carregamento (em caminhao) de marcantes de carga solta a partir do armazem, vinculando os marcantes do estoque a uma Ordem de Carregamento (OC) ja autorizada para um veiculo (placa cavalo + carreta). Permite tambem cancelar um carregamento bipado por engano e visualizar onde a carga estava no armazem.
- Usuario-alvo: operadores do armazem (Coletor) que conduzem o veiculo ate o armazem e bipam os marcantes para finalizar a saida da carga solta. Sistema executavel `ColetorArm.exe`.

## Acoes do usuario
- Acao principal 1: **Selecionar veiculo** no combo de placas para listar as Ordens de Carregamento (OCs) abertas para aquele cavalo (`PLACA_C`).
- Acao principal 2: **Bipar/informar marcante** no campo `MARCANTE` (9 digitos, IMEMode disabled). O sistema carrega automaticamente o lote, quantidade, embalagem, posicao no armazem e ids tecnicos (CS, CS_YARD, ARMAZEM).
- Acao principal 3: **`[F4] CARREGA ITEM`** - confirma o carregamento daquele marcante na OC selecionada (carregamento total ou parcial conforme `QTDE` informada x `VOLUMES` originais). Atualiza a posicao para `YARD='CAM'` (em caminhao).
- Acao principal 4: **`[F5] CANCELA`** - desfaz o carregamento de um marcante ja em `'CAM'`, retornando a posicao via insert de novo registro com `YARD='CANCC'` e zerando a placa do marcante.
- Acoes secundarias:
  - **`[F1] LOCAIS`** - abre painel sobreposto (`frmLocais`) listando os marcantes ainda no armazem (descricao do armazem + posicao + qtde) para a OC selecionada (`AUTONUMCS`). Util para o operador localizar a carga.
  - **`[F2] DOC`** - abre executavel externo `ConsultaDocColetor.exe` passando o BL (LTL) ou o agendamento (`MAX(autonum)` em `TB_AG_CS`) para visualizar documentos vinculados.
  - **`[F3] LIMPAR`** - limpa os campos do marcante (qtde, lote, local, armazem, marcante) e devolve foco ao campo marcante.
  - **`[...]` Refresh** - recarrega o combo de veiculos.
  - **`ESC`** - sai da tela (ou fecha o painel `frmLocais` quando aberto).

## Fluxo funcional
1. O usuario abre a tela pelo menu Principal do Coletor Armazem (acao "Carregamento Carga Solta"). Permissao validada por `Valida_Acesso_Botao`.
2. O combo `Veiculo` e populado com placas distintas em `VW_CAM_CARREGAMENTO` filtrando pelo `PATIO` do operador (`Col_Patio`).
3. O usuario seleciona o veiculo (placa cavalo + placa carreta + modelo). O sistema lista as OCs do veiculo na tabela `ORDENS`, mostrando `#ORDEM, LOTE, ITEM, QTDE_CARREGADA/QUANTIDADE, EMBALAGEM` (somatorio de marcantes ja em `CAM` para a placa selecionada ou ainda nulos).
4. O usuario seleciona uma OC. O sistema lista os marcantes ja carregados em `'CAM'` para aquela OC na tabela `CARREGAMENTOS`.
5. (Opcional) O usuario aciona `F1 LOCAIS` para visualizar onde os marcantes restantes estao no armazem.
6. (Opcional) O usuario aciona `F2 DOC` para abrir documentos do BL/agendamento vinculado a OC.
7. O usuario bipa um marcante. O sistema busca os dados (`Carrega_Dados`) e preenche QTDE, LOTE, LOCAL, ARMAZEM, e ids ocultos (`AUTONUMCS`, `AUTONUM_CS_YARD`).
8. O usuario aciona `F4 CARREGA ITEM`. O sistema valida (ver Regras) e, se OK, executa transacao para mover o marcante de `armazem -> 'CAM'`, registra historico e atualiza a placa do marcante.
9. O sistema recarrega `ORDENS` (atualiza `QTDE_CARREGADA`) e `CARREGAMENTOS` (acrescenta o marcante recem-bipado), reaplica selecao na OC anterior e devolve o foco ao campo marcante.
10. Caso o operador erre, ele bipa o marcante novamente (com `LOCAL='CAM'`) e aciona `F5 CANCELA`. O sistema reverte o lancamento (zera o registro `'CAM'`, insere `'CANCC'`, anula `PLACA_C`) e atualiza as listas.
11. Para sair, `ESC`. Se o painel `LOCAIS` estiver aberto, `ESC` apenas o fecha.

## Regras de uso percebidas
- Obrigatorio:
  - Selecionar um veiculo no combo (a tela depende disso para filtrar OCs).
  - Selecionar uma OC na lista `ORDENS` antes de bipar.
  - Marcante informado/bipado em campo de 9 digitos.
  - Marcante deve estar associado a uma `TB_CARGA_SOLTA` (carga solta de importacao) - via `TB_MARCANTES.AUTONUM_CARGA`.
  - O `AUTONUMCS` carregado pelo marcante deve ser igual ao `AUTONUMCS` da OC selecionada (caso contrario, "Carga Divergente da Ordem de carregamento").
  - O marcante nao pode ja estar em `'CAM'` para carregar (mensagem "Carga ja carregada"). Para `F5`, e o oposto: precisa estar em `'CAM'`.
  - Quantidade `txtQtde` <= quantidade original (`txtQtde.Tag` = `VOLUMES`); senao "Quantidade indisponivel".
  - Funcao `FC_VALIDA_SAIDA_CARGA(BL, ORDEM, AUTONUMCS)` deve retornar `'OK'` para liberar carregamento; qualquer outro retorno e exibido como critica.
- Opcional:
  - Painel `LOCAIS` e `DOC` sao auxiliares e nao sao requisitos para `F4`.
  - Quantidade pode ser ajustada manualmente (parcial) antes do `F4` (ainda assim limitada por `Tag`).
- Restricoes:
  - Quando `Wflag_carregamento_marcante_bip <> 0`, marcante so pode ser inserido por leitor de codigo de barras (validacao por `DateDiff` entre Enter e digitacao). Mensagem "Campo nao permite digitacao manual".
  - Quando `Wflag_carregamento_marcante_bip = 1`, apos validar marcante, o campo e desabilitado ate proximo ciclo (forca `Limpa` para reabilitar).
  - Filtro de pais/patio (`Col_Patio`): operadores so veem placas/OCs do seu patio (com regra especial `Col_Patio=1` -> `PATIO IN (1,7)`).

## Mensagens funcionais
- Sucesso (em `txtCritica` com `ForeColor = vbGreen`):
  - "ITEM CARREGADO" (carregamento total).
  - "ITEM PARCIALMENTE CARREGADO" (qtde menor que original).
  - "ITEM CANCELADO" (cancelamento).
- Alerta/Erro (em `txtCritica` com `ForeColor = vbRed`):
  - "Escolha uma carga" - `F2` sem `txtLote` valido.
  - "Ordem de carregamento nao selecionada" - `F4` sem OC selecionada.
  - "Marcante nao informado/invalido" - `F4` ou `F5` sem marcante.
  - "Carga Divergente da Ordem de carregamento" - `F4` com `AUTONUMCS` diferente do da OC.
  - "Carga ja carregada" - `F4` com `LOCAL='CAM'`.
  - "Carga ainda nao carregada" - `F5` com `LOCAL <> 'CAM'`.
  - "Quantidade indisponivel" - `F4` com `txtQtde > txtQtde.Tag`.
  - "Lote nao encontrado" - falha ao recuperar `TB_CARGA_SOLTA` por `AUTONUMCS`.
  - Texto livre vindo da function `FC_VALIDA_SAIDA_CARGA` (ex.: bloqueios SISCARGA, GLME, free-time, perfil, BL, etc.).
- Modal (`MsgBox`):
  - "Campo nao permite digitacao manual" - usuario configurado para bipagem digitou manualmente.

## Pontos de atencao
- Dependencias operacionais:
  - Operador precisa estar logado, com vinculo a um `PATIO` valido em `TB_CAD_USUARIOS`.
  - Operador precisa de permissao `COLETOR_ARM` para `COMMAND1(7)` em `TB_SYS_FUNCOES`/`TB_SYS_GRP_PERMISSOES`.
  - O cadastro `TB_CAD_USUARIOS.flag_carregamento_marcante_bip` impacta diretamente a UX (digitar vs bipar).
  - Veiculo precisa estar agendado (gerar linhas em `VW_CAM_CARREGAMENTO`) para aparecer no combo. Se nao aparece, nada se faz na tela.
  - Funcao `FC_VALIDA_SAIDA_CARGA` precisa estar saudavel; sua indisponibilidade trava todo o fluxo.
  - `ConsultaDocColetor.exe` precisa estar instalado em `App.Path\ConsultaDocumentos`. Em ambiente web, esta integracao precisa ser repensada.
- Casos de borda:
  - Marcante com `PLACA_C` ja preenchida (e diferente do veiculo selecionado): nao aparece no `Carrega_LvOrdem` (`PLACA_C IS NULL OR PLACA_C='<placa>'`). O usuario nao tem visibilidade de que foi "puxado" por outro veiculo.
  - Carregamento parcial: a tela atualiza `VOLUMES` do marcante para o valor parcial e atualiza `tb_temp_marcante.QTDE`. O remanescente segue para a tela de "Associacao de Marcantes" (?) - hipotese, requer confirmacao.
  - Recuperacao do `AUTONUM` recem-inserido em `TB_CARGA_SOLTA_YARD` via `MAX(AUTONUM)` filtrado por `AUTONUM_CS`: pode haver risco de concorrencia (dois operadores na mesma OC).
  - Tela permite `F4` mesmo quando o `lvOC.SelectedItem` virou inconsistente apos refresh; o codigo trata via `If Not (Me.lvOC.SelectedItem Is Nothing)`.
  - Combo de veiculos contem `PLACA_C + ' ' + PLACA_CARRETA + ' - ' + ISNULL(MODELO,'')`. O sistema usa apenas `Left(text, 8)` para extrair a placa do cavalo - portanto a placa do cavalo precisa caber em 8 caracteres (formato Mercosul `ABC1D23` ou `ABC-1234`).
  - `lvOC` em VB6 quando `KeyUp` muda selecao: o `Carrega_LvCarregamentos` reflete a nova selecao - util para teclado.
  - `Carrega_Dados` retorna `Limpa` quando o marcante nao tem `AUTONUM_CARGA` valido em `TB_CARGA_SOLTA` (cargas de exportacao ou marcantes invalidos sao filtrados implicitamente pelo INNER JOIN).
