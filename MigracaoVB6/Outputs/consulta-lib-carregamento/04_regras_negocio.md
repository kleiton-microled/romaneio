# Regras de Negocio - Consulta Lib. Carregamento

## Regras identificadas por acao

### Filtrar
- Regra (Regra de validacao): pelo menos um caminho de filtro deve ser informado.
- Evidencia: `cmdFiltra_Click` bloqueia quando protocolo/placa/BL ausentes e quando placa sem CNTR/BL.

- Regra (Regra de negocio): usuario so pode operar cargas do proprio patio (com excecao de combinacao patio 1/7).
- Evidencia: validacao `Col_Patio` versus `Rst!Patio` com mensagem de divergencia.

- Regra (Regra de interface): protocolo e formatado automaticamente para padrao `######/####`.
- Evidencia: uso de `ObtemNumero`, `Right`, `Mid` antes das consultas.

### Incluir
- Regra: nao aplicavel diretamente (tela nao possui inclusao manual de itens); inclusao ocorre no registro/liberacao.
- Evidencia: ausencia de fluxo "incluir" explicito no form.

### Remover
- Regra: nao aplicavel diretamente (sem acao de remocao na tela).
- Evidencia: ausencia de botao/remocao de registros no form.

### Salvar/Persistir
- Regra (Regra de negocio): nao registrar se houver bloqueio de BL ou conteiner.
- Evidencia: `cmdRegistra_Click` verifica labels `lblbloqueioBL` e `lblbloqueioCntr`.

- Regra (Regra de negocio): nao registrar se houver pendencia GLME.
- Evidencia: consultas em `TB_registro_saida_cs/cntr` cruzadas com `TB_etapas_carga` e bloqueio com mensagem.

- Regra (Regra de negocio): validar pagamento/free-time antes do registro.
- Evidencia: chamadas `VerificaFormaPagamento`, `VerificaPendenteVista` e consulta `TB_GR_BL`.

- Regra (Regra de negocio): respeitar janela de agendamento; fora da janela requer permissao especial.
- Evidencia: calculo `DateDiff` com `TB_GD_RESERVA.JANELA` e validacao em `tb_cad_usuarios.flag_libera_janela`.

- Regra (Regra tecnica): atualiza tipo de veiculo em `TB_AG_VEICULOS` no ato de registrar.
- Evidencia: `UPDATE TB_AG_VEICULOS SET ID_TIPO_CAMINHAO=...`.

- Regra (Regra de negocio): consulta/liberacao depende de flags de patio.
- Evidencia: `FLAG_CONSULTA_LIBERACAO_COLETOR` e `FLAG_CONSULTA_LIBERACAO`.

- Regra (Regra de negocio): em falha SISCARGA/SEFAZ, pode haver continuidade condicionada (conforme perfil e tipo documental).
- Evidencia: fluxo com `Liberar_Bloqueado`, `EhDTA`, mensagens de confirmacao e chamadas a `atualizaLiberacao`.

### Sair/Navegar
- Regra (Regra de interface): `ESC` encerra tela.
- Evidencia: `Form_KeyDown`.

## Validacoes
- Campo: protocolo/placa/BL/CNTR
  - Condicao: pelo menos uma combinacao valida.
  - Mensagem: "Informe o BL, placa ou protocolo" / "Informe a placa" / "Informe o conteiner ou o BL".

- Campo: ordem de carregamento
  - Condicao: deve existir ordem valida para continuar.
  - Mensagem: "Registro de Saida nao localizado ...".

- Campo: patio
  - Condicao: patio da carga deve ser compativel com usuario.
  - Mensagem: "Patio divergente do usuario.".

- Campo: bloqueios/GLME
  - Condicao: sem bloqueio de BL/CNTR e sem pendencia GLME.
  - Mensagem: "Existe Bloqueio..." / "Atenção, Existe Carga sem Liberacao GLME .".

## Regras de consistencia
- Quantidade:
  - Sem regra quantitativa de composicao no form; consistencia principal e de elegibilidade para liberar saida.
- Integridade referencial:
  - consulta cruza multiplas entidades (ordem, registro saida, BL, patio, motorista, transportadora) e depende de links consistentes.
- Concorrencia:
  - nao ha evidencia de lock/transacao explicita no form VB6; comportamento depende das funcoes de dominio chamadas.

## Efeitos colaterais
- Atualizacoes indiretas:
  - `atualizaLiberacao` e `registraSaidaComex` alteram estado operacional fora do escopo visual imediato.
  - atualizacao de `TB_ORDEM_CARREGAMENTO` (usuario, periodo/janela, flags).
- Reuso de tabela temporaria:
  - nao evidenciado uso de tabela temporaria local do form para staging.
- Dependencias externas:
  - integracoes SISCARGA/SEFAZ.
  - executavel externo de documentos.

## Hipoteses e duvidas
- Hipotese:
  - parte principal das regras fiscais/aduaneiras esta centralizada em funcoes externas (`validaSaidaComex`, `registraSaidaComex`, `atualizaLiberacao`).
- Evidencia parcial:
  - form apenas prepara contexto e consome `msgdoc`/`msgIcms`, sem implementar logica completa localmente.
- Acao de validacao:
  - revisar implementacao dessas funcoes no modulo legado compartilhado antes de codificar repositorio MVC.

- Hipotese:
  - fluxo do botao "Documentos" e suporte operacional, nao pre-condicao tecnica obrigatoria da liberacao.
- Evidencia parcial:
  - botao separado e executavel externo.
- Acao de validacao:
  - confirmar com usuario funcional se primeira entrega MVC deve incluir abertura externa.
