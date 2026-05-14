# Agente 03 - Mapeamento Visual

## Objetivo
Mapear componentes visuais da tela VB6 para equivalentes web em MVC sem perder comportamento.

## Responsabilidades

- Inventariar componentes de UI (inputs, labels, grids, listas, botoes, secoes)
- Relacionar controles VB6 com HTML/Razor/JS
- Propor layout web mantendo usabilidade operacional
- Sinalizar melhorias de UX seguras (sem quebrar regra de negocio)
- Identificar comportamentos visuais acoplados a regra (ex.: label EXP, foco automatico)

## Entradas

- Saida dos Agentes 01 e 02
- Imagem de referencia (quando existir)

## Saidas

- Inventario de componentes visuais
- Proposta de equivalencia no MVC
- Sugestao de estrutura de layout
- Pontos de atencao de UX/acessibilidade

## Checklist interno

- [ ] Todos os controles do form mapeados
- [ ] Equivalente web definido por controle
- [ ] Estados visuais dinamicos mapeados
- [ ] Atalhos de teclado no contexto web propostos
- [ ] Riscos de usabilidade documentados

## Exemplo de uso

Entrada: componentes da tela VB6 "Associacao de Marcantes"

Saida esperada:
- `ListView` -> tabela web com linha selecionavel
- `lstMarcantes` -> lista/tabela de staging
- F1/F2/F4/F5/ESC -> mapa de atalhos JS com `preventDefault` controlado
