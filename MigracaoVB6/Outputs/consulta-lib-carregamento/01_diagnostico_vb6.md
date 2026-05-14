# Diagnostico VB6 - Consulta Lib. Carregamento

## Identificacao da tela
- Nome: Consulta - Lib. Carregamento (caption do menu principal).
- Modulo: Coletor (VB6).
- Form principal: `frmLibRegistro` (fonte principal em `frmLibRegistroS.frm`; existe variante em `frmLibRegistro.frm`).
- Caminho: `ProjetoReferencia/Coletor`.

## Arquivos localizados
- `.frm`:
  - `ProjetoReferencia/Coletor/PrincipalS.frm` (entrada pelo menu `Command1(Index=9)`).
  - `ProjetoReferencia/Coletor/frmLibRegistroS.frm` (tela principal analisada).
  - `ProjetoReferencia/Coletor/frmLibRegistro.frm` (versao alternativa/legada simplificada).
- `.frx`:
  - `ProjetoReferencia/Coletor/frmLibRegistroS.frx`.
- `.bas`:
  - `ProjetoReferencia/Coletor/Microled.bas` (globais de conexao, utilitarios e estado).
  - `ProjetoReferencia/Coletor/mdlColetor.bas` (globais operacionais como patio).
  - `ProjetoReferencia/Coletor/GERENCIAL.BAS` (modulo grande com regras compartilhadas; chamado indiretamente pelo form).
- `.cls`:
  - Nao identificados com evidencia direta para esta tela.
- Outros:
  - `ProjetoReferencia/Coletor/Coletor.vbp` (estrutura do projeto).
  - `ProjetoReferencia/Coletor/ConsultaDocumentos/ConsultaDocColetor.exe` (acionado por botao "Documentos").

## Eventos mapeados
- Form_Load:
  - Carrega `cbTipoVeiculo` via query em `TB_TIPOS_CAMINHAO`.
- Form_KeyDown/KeyPress:
  - `ESC` fecha o form (`Unload Me`).
- Clicks de botoes:
  - `cmdFiltra_Click`: valida filtros, consulta dados, carrega lista de carga, status e bloqueios.
  - `cmdRegistra_Click`: valida regras, tenta liberar/registrar saida (SISCARGA/SEFAZ), atualiza situacao.
  - `cmdLimpar_Click`: limpa todos os campos e estados de tela.
  - `Command1_Click` (Documentos): abre executavel externo `ConsultaDocColetor.exe`.
- Eventos de lista/grid:
  - Nao ha grid rico; `listaCarga` e `lstSimultaneos` sao preenchidos programaticamente por consultas.
- Eventos de foco:
  - SetFocus explicito apos limpar/sucesso para `mskPlaca`.

## SQLs e persistencia
- Consultas (principais):
  - Busca principal por protocolo/placa/BL/CNTR em `VW_AGENDA_CS`, `TB_BL`, `TB_PATIOS`, `TB_CNTR_BL`, `TB_AMR_CNTR_BL`, `TB_GD_RESERVA`, `TB_AG_VEICULOS`, `TB_REGISTRO_SAIDA_CS`, `TB_REGISTRO_SAIDA_CNTR`, `TB_ORDEM_CARREGAMENTO`, `TB_CARGA_SOLTA`.
  - Carregamento da lista de cargas (CS/CNTR) para exibicao.
  - Consulta de localizacao: `vw_invent_armazem_item`.
  - Consulta de protocolos simultaneos por placa+periodo em `VW_AGENDA_CS`.
  - Validacoes de bloqueio/GLME/free-time/pagamento/perfil.
- Inserts:
  - Nao ha insert direto evidente no form; registro final parece encapsulado em funcoes de dominio (`registraSaidaComex`, `atualizaLiberacao`).
- Updates:
  - `UPDATE TB_AG_VEICULOS SET ID_TIPO_CAMINHAO=...`.
  - `UPDATE TB_ORDEM_CARREGAMENTO` para usuario/perfil de janela/flags.
  - Outros updates indiretos via funcoes chamadas.
- Deletes:
  - Nao evidenciado diretamente no form.
- Views/Tabelas/Functions:
  - Views: `VW_AGENDA_CS`, `vw_invent_armazem_item`.
  - Tabelas: `TB_BL`, `TB_PATIOS`, `TB_CNTR_BL`, `TB_AMR_CNTR_BL`, `TB_CARGA_SOLTA`, `TB_REGISTRO_SAIDA_CS`, `TB_REGISTRO_SAIDA_CNTR`, `TB_ORDEM_CARREGAMENTO`, `TB_ETAPAS_CARGA`, `TB_TIPOS_CAMINHAO`, `TB_SOLICITACAO_LTL`, `TB_MOTORISTAS`, `TB_CAD_TRANSPORTADORAS`, `TB_CAD_USUARIOS`.
  - Funcoes/rotinas chamadas: `validaEntregaDoc`, `validaSaidaComex`, `registraSaidaComex`, `atualizaLiberacao`, `VerificaFormaPagamento`, `VerificaPendenteVista`, `ObtemNumero`, `Nnull`, `PRSet`, `Executa`.

## Dependencias tecnicas
- Variaveis globais:
  - `Banco_Sgipa`, `Banco_Operador`, `msgIcms`, `msgdoc`, `db`, `ConnectionString`, `Cod_Usuario` (de `Microled.bas`).
  - `Col_Patio` (de `mdlColetor.bas`).
- Funcoes utilitarias:
  - `ObtemNumero`, `Nnull`, `PRSet` e demais helpers de acesso a dados.
- Formularios auxiliares:
  - `Principal`/`PrincipalS` (menu de entrada).
- Controle de permissao:
  - Validacao de patio (`Col_Patio` x patio da carga).
  - Validacao de usuario para liberar fora de janela (`tb_cad_usuarios.flag_libera_janela`).
  - Parametrizacao por patio (`FLAG_CONSULTA_LIBERACAO_COLETOR` e `FLAG_CONSULTA_LIBERACAO`).

## Atalhos e automatismos
- Atalhos:
  - `ESC` para sair.
- Foco automatico:
  - Ao limpar/sucesso, foco retorna para `mskPlaca`.
- Comportamentos implicitos:
  - Mascaras de entrada para protocolo/placa/cntr.
  - Autoformatacao do protocolo (`######/####`).
  - Exibicao condicional de labels de bloqueio e GLME.
  - Botao registrar habilita/desabilita conforme validacoes preliminares.

## Pontos criticos
- Riscos:
  - SQL dinamico concatenado (risco de erro e manutencao).
  - Forte acoplamento a globais e a banco diretamente no form.
  - Fluxo com muitos `Exit Sub` e mensagens condicionais dificulta rastreabilidade.
- Inconsistencias:
  - Codificacao de caracteres acentuados aparece corrompida nos fontes.
  - Nomes semelhantes para telas (`frmLibRegistro.frm` vs `frmLibRegistroS.frm`) podem induzir erro de mapeamento.
- Hipoteses:
  - Hipotese: a versao em uso operacional para "Consulta - Lib. Carregamento" e `frmLibRegistroS.frm` por conter filtros ampliados (placa/CNTR/BL), bloqueios e botao de documentos.
  - Hipotese: parte das regras de liberacao efetiva esta em modulos externos (`GERENCIAL.BAS`/shared), nao apenas no form.
