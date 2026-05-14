# Mapeamento Funcional - Consulta Lib. Carregamento

## Finalidade da tela
- Objetivo de negocio: consultar a situacao de liberacao de carregamento (SISCARGA/SEFAZ), exibir dados da carga e permitir o registro/liberacao da saida quando autorizado.
- Usuario-alvo: operador de registro/coletor com permissao no patio e permissao operacional para liberar saida.

## Acoes do usuario
- Acao principal 1: filtrar carga por protocolo, ou por placa com apoio de CNTR/BL.
- Acao principal 2: registrar/liberar a saida da carga apos validacoes de negocio.
- Acoes secundarias:
  - limpar tela e reiniciar pesquisa.
  - consultar documentos relacionados via executavel externo.
  - visualizar protocolos simultaneos.
  - visualizar status de bloqueio e alerta GLME.

## Fluxo funcional
1. Usuario informa criterios (`protocolo` ou `placa` + `CNTR/BL`) e aciona **Filtrar**.
2. Sistema valida entradas minimas, consulta base, exibe dados da carga (periodo, lote, BL, transportadora, motorista, placas, tipo veiculo, localizacao, lista de cargas e status SISCARGA/SEFAZ).
3. Sistema verifica patio do usuario, bloqueios e parametrizacao de consulta/liberacao.
4. Usuario aciona **Registrar**; sistema executa regras adicionais (GLME, pagamento/free-time, janela, permissao especial, retorno SISCARGA/SEFAZ) e conclui com sucesso ou bloqueio/mensagem.
5. Em sucesso, tela limpa e retorna foco para nova operacao.

## Regras de uso percebidas
- Obrigatorio:
  - informar pelo menos um caminho valido de filtro.
  - sem protocolo, exige placa e pelo menos CNTR ou BL.
- Opcional:
  - quando protocolo existe, placa/CNTR/BL atuam como refinadores.
  - ajuste de tipo de veiculo antes do registro.
- Restricoes:
  - usuario deve operar no patio da carga.
  - bloqueio de BL/CNTR impede registro.
  - pendencias de GLME/faturamento/free-time podem impedir registro.
  - parametros de patio podem ocultar/desabilitar liberacao.

## Mensagens funcionais
- Sucesso:
  - "Registrado com sucesso"
  - "Saida Registrada com sucesso no SISCARGA/SEFAZ!"
- Alerta:
  - "Patio divergente do usuario."
  - "Atenção, Existe Carga sem Liberacao GLME."
  - avisos de janela fora de periodo com confirmacao.
- Erro:
  - "Informe o BL, placa ou protocolo"
  - "Informe a placa"
  - "Informe o conteiner ou o BL"
  - "Registro de Saida nao localizado..."
  - "Nao foi possivel registrar a saida da carga!"

## Pontos de atencao
- Dependencias operacionais:
  - disponibilidade de integracoes e retorno de SISCARGA/SEFAZ.
  - consistencia das tabelas de registro de saida e ordem de carregamento.
  - executavel externo `ConsultaDocColetor.exe` para documentos.
- Casos de borda:
  - consulta retornando agenda sem ordem de carregamento valida.
  - retorno externo vazio (msgdoc/msgIcms vazios) com fluxo de confirmacao manual.
  - cargas LTL com mensagem especifica de liberacao total do protocolo.
