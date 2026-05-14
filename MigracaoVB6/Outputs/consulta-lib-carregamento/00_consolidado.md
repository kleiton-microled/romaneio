# Consolidado - Consulta Lib. Carregamento

## Resumo executivo
- Tela: Consulta - Lib. Carregamento (`frmLibRegistro` em `frmLibRegistroS.frm`).
- Objetivo: consultar e registrar liberacao de saida de cargas com validacoes de patio, bloqueios e integracoes SISCARGA/SEFAZ.
- Status do estudo: concluido (fase documental), sem implementacao MVC.

## Principais achados
- Achado 1:
  - fluxo principal depende de filtros combinados (protocolo/placa/CNTR/BL), com regras de obrigatoriedade diferentes por contexto.
- Achado 2:
  - regra de registro e extensa, com validacoes de bloqueio, GLME, pagamento/free-time, janela e permissao de usuario.
- Achado 3:
  - comportamento operacional depende de funcoes compartilhadas externas ao form (`validaSaidaComex`, `registraSaidaComex`, `atualizaLiberacao`).

## Regras criticas
- Regra:
  - sem dados minimos de filtro, a consulta nao deve prosseguir.
- Regra:
  - sem compatibilidade de patio, bloqueios resolvidos e validacoes de liberacao aprovadas, o registro deve ser impedido.
- Regra:
  - parametricas de patio (`FLAG_CONSULTA_LIBERACAO_*`) controlam visibilidade e permissao da acao registrar.

## Decisoes de migracao
- Decisao:
  - criar modulo MVC dedicado para a tela (`ConsultaLiberacaoCarregamentoController` + repositorio + view + js).
- Justificativa:
  - concentracao de regras de negocio criticas e necessidade de isolamento para manutencao/teste.

- Decisao:
  - usar Dapper com SQL parametrizado, eliminando concatenacao direta do VB6.
- Justificativa:
  - seguranca e reducao de risco tecnico sem alterar semantica funcional.

## Criterios de aceite da migracao
- Filtros:
  - deve consultar por protocolo.
  - sem protocolo, deve aceitar placa combinada com CNTR ou BL.
- Colunas/campos exibidos:
  - periodo, lote, BL, transportadora, motorista, placas, localizacao, tipo veiculo, status SISCARGA e SEFAZ.
- Acoes disponiveis:
  - filtrar, registrar (quando permitido), limpar, sair.
- Validacoes obrigatorias:
  - patio divergente, bloqueio BL/CNTR, pendencia GLME, pendencias de pagamento/free-time, janela/permissao e parametros de patio.
- Preservacao de comportamento VB6:
  - formatacao de protocolo, mascaras de entrada, mensagens de negocio equivalentes e foco operacional apos limpeza/sucesso.

## Pendencias de validacao
- Pendencia:
  - confirmar implementacao detalhada e semantica final das funcoes compartilhadas de liberacao (`validaSaidaComex`, `registraSaidaComex`, `atualizaLiberacao`).
- Responsavel:
  - time tecnico (desenvolvimento + DBA).

- Pendencia:
  - definir estrategia web para acao "Documentos" (executavel externo no VB6).
- Responsavel:
  - negocio + arquitetura.

## Objetos de banco envolvidos (levantados)
- Views:
  - `VW_AGENDA_CS`, `vw_invent_armazem_item`.
- Tabelas principais:
  - `TB_BL`, `TB_PATIOS`, `TB_CNTR_BL`, `TB_AMR_CNTR_BL`, `TB_CARGA_SOLTA`, `TB_REGISTRO_SAIDA_CS`, `TB_REGISTRO_SAIDA_CNTR`, `TB_ORDEM_CARREGAMENTO`, `TB_ETAPAS_CARGA`, `TB_GD_RESERVA`, `TB_TIPOS_CAMINHAO`, `TB_AG_VEICULOS`, `TB_SOLICITACAO_LTL`, `TB_CAD_USUARIOS`, `TB_MOTORISTAS`, `TB_CAD_TRANSPORTADORAS`, `TB_GR_BL`.

## Riscos e lacunas
- Risco:
  - logica distribuida entre form e modulos externos pode gerar perda de regra na migracao se nao houver rastreio completo.
- Risco:
  - alto volume de condicionais com mensagens pode causar divergencia de UX e comportamento.
- Lacuna:
  - nao foi localizada evidencia completa do corpo das funcoes de integracao dentro deste estudo (chamadas identificadas, implementacao detalhada pendente de aprofundamento).

## Escopo da implementacao (fase atual)
- Incluido:
  - diagnostico tecnico, mapeamento funcional/visual, catalogo de regras, desenho MVC aderente ao Romaneio e plano de implementacao.
- Fora de escopo:
  - implementacao da tela MVC, alteracoes de banco e testes automatizados.

## Proximas etapas
1. Validar pendencias em aberto com funcional e DBA.
2. Iniciar implementacao pela Etapa 1 (`06_plano_implementacao.md`).
3. Homologar comportamento da primeira versao MVC contra roteiro de aceite.

## Prontidao para implementacao
- Nivel de prontidao: Medio-Alto.
- Condicoes para inicio seguro:
  - confirmar regras internas das funcoes compartilhadas de liberacao.
  - decidir tratamento web para integracao de documentos.
