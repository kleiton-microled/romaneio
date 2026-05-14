# Orquestrador - Migracao de Tela VB6 para MVC

## Finalidade
Definir a ordem, entrada, saida e acoplamento entre agentes para migrar qualquer tela VB6 do `ProjetoReferencia` para o `Romaneio`.

## Ordem de execucao padrao

1. **A01 Leitura VB6**
2. **A02 Mapeamento funcional**
3. **A03 Mapeamento visual**
4. **A04 Regras de negocio**
5. **A05 Desenho MVC**
6. **A06 Plano de implementacao**

## Contrato entre agentes

### A01 -> A02
- Entrega: arquivos, eventos, SQLs, dependencias
- Uso: base para entender "o que a tela faz"

### A02 + A01 -> A03
- Entrega: funcoes de negocio + controles existentes
- Uso: equivalencia visual e organizacao da UI web

### A01 + A02 + A03 -> A04
- Entrega: comportamento tecnico e funcional combinado
- Uso: consolidar regras reais, validacoes e hipoteses

### A04 + estrutura Romaneio -> A05
- Entrega: regras e restricoes finais
- Uso: desenhar solucao aderente ao projeto existente

### A05 -> A06
- Entrega: arquitetura alvo e componentes
- Uso: quebrar em fases executaveis e checklist final

## Entradas globais do processo

- Nome da tela
- Caminho da tela no VB6 (quando conhecido)
- Imagem/referencia visual (opcional)
- Restricoes de negocio/prazo

## Saidas globais do processo

- Diagnostico tecnico legado
- Documento funcional
- Mapeamento visual
- Catalogo de regras de negocio
- Proposta arquitetural MVC
- Plano de implementacao + checklist

## Forma de consolidacao

Ao final, gerar `00_consolidado.md` com:

- Resumo executivo da tela
- Decisoes confirmadas
- Hipoteses em aberto
- Itens bloqueantes
- Escopo da primeira implementacao

## Gatilhos de retorno (loop)

Voltar para agente anterior quando:

- A03 detectar controle visual sem fonte no A01
- A04 detectar regra sem evidencia tecnica
- A05 detectar dependencia nao mapeada no A01
- A06 detectar lacuna impeditiva de execucao

## Criterio de pronto da fase de estudo

- Todos os templates 01..06 preenchidos
- Hipoteses explicitadas
- Sem item critico sem dono (responsavel por validacao)
