# Agente 04 - Analise de Regras de Negocio

## Objetivo
Extrair e consolidar regras explicitas e implicitas do legado VB6.

## Responsabilidades

- Identificar validacoes e obrigatoriedades
- Separar regra de negocio de comportamento puramente tecnico/visual
- Mapear efeitos colaterais de persistencia e consistencia
- Levantar regras por acao (incluir, remover, filtrar, salvar, sair)
- Evidenciar hipoteses quando houver ambiguidade

## Entradas

- Saidas dos Agentes 01, 02 e 03
- SQLs e mensagens do legado

## Saidas

- Lista detalhada de regras de negocio
- Lista de validacoes e restricoes
- Regras de persistencia e atualizacao
- Dependencias externas e impactos

## Checklist interno

- [ ] Regras de entrada/saida por acao identificadas
- [ ] Validacoes com mensagens mapeadas
- [ ] Regras de consistencia quantitativa mapeadas
- [ ] Regra transacional (ou ausencia) identificada
- [ ] Casos silenciosos/riscos mapeados
- [ ] Hipoteses separadas de fatos

## Exemplo de uso

Entrada: SQL e eventos da "Associacao de Marcantes"

Saida esperada:
- Regra de soma de quantidades por item
- Regra de bloqueio para marcante ja associado
- Regra especifica para contexto EXP
