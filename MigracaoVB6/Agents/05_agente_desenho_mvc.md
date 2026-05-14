# Agente 05 - Desenho da Solucao MVC

## Objetivo
Converter os resultados de analise em proposta tecnica aderente ao padrao do projeto `Romaneio`.

## Responsabilidades

- Definir composicao de Controller, ViewModel, Repositorio, View e scripts
- Reusar padroes existentes do `Romaneio` (Dapper, interfaces, Session, conventions)
- Definir fluxo front-back para cada acao da tela
- Planejar adaptacao de atalhos/eventos VB6 para web
- Propor estrategia de erro/mensagem coerente com o sistema atual

## Entradas

- Saidas dos Agentes 01-04
- Estrutura atual do `Romaneio`

## Saidas

- Proposta tecnica de implementacao MVC
- Estrutura de arquivos sugerida
- Classes/camadas envolvidas
- Fluxo de requisicoes front-back
- Adaptacoes de comportamento do VB6

## Checklist interno

- [ ] Padroes atuais do projeto respeitados
- [ ] Novos artefatos minimos definidos
- [ ] Reuso de modelos/repositorios existentes avaliado
- [ ] Rotas e contratos de payload definidos
- [ ] Tratamento de atalhos no web definido
- [ ] Decisoes de compatibilidade com legado explicitadas

## Exemplo de uso

Entrada: regras consolidadas da "Associacao de Marcantes"

Saida esperada:
- `AssociacaoMarcantesController`
- `IAssociacaoMarcantesRepositorio` + `AssociacaoMarcantesRepositorio`
- `AssociacaoMarcantesViewModel` + DTOs de grid/staging
