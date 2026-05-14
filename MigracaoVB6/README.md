# Framework Interno de Migracao VB6 -> MVC

Este pacote organiza a migracao de telas VB6 (em `ProjetoReferencia`) para o projeto MVC C# `Romaneio` com um fluxo padronizado e reutilizavel.

## Objetivo

- Evitar analises superficiais
- Padronizar o diagnostico por tela
- Reduzir risco de perda de regras implicitas do VB6
- Entregar insumos prontos para implementacao MVC

## Componentes

- `Agents/`: definicao dos 6 agentes internos com checklist e entradas/saidas
- `Playbooks/`: orquestrador do processo ponta-a-ponta
- `Templates/`: modelos markdown para preencher em cada tela
- `Outputs/`: pasta alvo para guardar saidas reais de cada migracao
- `Examples/`: exemplo aplicado (Associacao de Marcantes)

## Fluxo resumido

1. Agente de leitura do VB6
2. Agente de mapeamento funcional
3. Agente de mapeamento visual
4. Agente de analise de regras de negocio
5. Agente de desenho da solucao MVC
6. Agente de plano de implementacao

## Convencao recomendada de artefatos

Criar uma pasta por tela em `Outputs/<nome-da-tela>/` com os templates preenchidos:

- `01_diagnostico_vb6.md`
- `02_mapeamento_funcional.md`
- `03_mapeamento_visual.md`
- `04_regras_negocio.md`
- `05_desenho_mvc.md`
- `06_plano_implementacao.md`
- `00_consolidado.md` (visao integrada final)
