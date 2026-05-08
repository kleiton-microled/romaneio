# Documentacao de Uso - Consulta Lib. Carregamento (MVC)

## Objetivo da tela
Permitir consultar dados da carga e registrar a liberacao de saida no contexto operacional de carregamento, preservando as regras principais do legado VB6.

## Quando utilizar
- Quando for necessario verificar se uma carga pode ser liberada.
- Quando for necessario validar status e bloqueios antes de registrar a saida.
- Quando a operacao precisar consultar por protocolo, placa, CNTR ou BL.

## Como acessar
- Pelo menu principal: **Consulta Lib. Carregamento**.
- Ou diretamente pela rota: `~/ConsultaLiberacaoCarregamento/Index`.

## Campos e componentes principais
- **Protocolo**: filtro por protocolo de agenda (padrao esperado `######/####`).
- **Placa**: filtro por placa do cavalo.
- **CNTR**: filtro complementar por conteiner.
- **BL**: filtro complementar por documento BL.
- **Periodo, Lote, N Documento, Transportadora, Motorista, Placas, Localizacao**: dados retornados da consulta.
- **Tipo Veiculo**: tipo de caminhao para registro.
- **Status Siscarga / Status Sefaz**: status de consulta/liberacao.
- **Cargas**: lista de cargas vinculadas ao registro encontrado.
- **Protocolos Simultaneos**: protocolos no mesmo contexto de placa/periodo.
- **Indicadores**:
  - `Bloq. BL`
  - `Bloq. Conteiner`
  - `GLME pendente`

## Fluxo recomendado de operacao
1. Informar os filtros:
   - protocolo; ou
   - placa (com CNTR ou BL quando nao houver protocolo).
2. Acionar **Filtrar**.
3. Conferir os dados da carga retornada.
4. Verificar os indicadores de bloqueio/GLME.
5. Ajustar **Tipo Veiculo** (quando aplicavel).
6. Se permitido, acionar **Registrar**.
7. Confirmar mensagem de sucesso e seguir para o proximo atendimento.

## Regras de validacao importantes
- Sem informar protocolo/placa/BL, a consulta e bloqueada.
- Sem protocolo, placa e obrigatoria.
- Sem protocolo, e necessario informar CNTR ou BL em conjunto com placa.
- Usuario so pode registrar carga com patio compativel.
- Com bloqueio de BL/CNTR, o registro e bloqueado.
- Com pendencia GLME, o registro e bloqueado.
- Parametrizacao de patio pode impedir liberacao.

## Mensagens comuns
- **Validacao de entrada**:
  - "Informe o BL, placa ou protocolo"
  - "Informe a placa"
  - "Informe o conteiner ou o BL"
- **Restricao operacional**:
  - "Patio divergente do usuario."
  - "Existe Bloqueio de BL"
  - "Existe Bloqueio de Conteiner"
  - "Atencao, Existe Carga sem Liberacao GLME ."
- **Resultado**:
  - "Registrado com sucesso"
  - "Nao foi possivel registrar a saida da carga!"

## Atalhos de teclado
- **ESC**: sair da tela.

> Observacao: no ambiente web, atalhos podem variar por navegador/politica local. Sempre use os botoes da interface como fallback.

## Botao Documentos
- A funcao existe na tela MVC como ponto de operacao.
- A integracao final com o executavel externo do legado (`ConsultaDocColetor.exe`) depende de definicao funcional/arquitetural.

## Boas praticas de uso
- Priorizar consulta por protocolo quando disponivel.
- Sempre revisar bloqueios e status antes de registrar.
- Em caso de erro, ajustar os filtros e repetir a consulta.
- Registrar observacao operacional quando houver impedimento por regra de negocio.

## Resultado esperado
Com dados validos e sem restricoes impeditivas, o registro da liberacao e concluido com sucesso, mantendo o fluxo operacional esperado para homologacao tecnica.
