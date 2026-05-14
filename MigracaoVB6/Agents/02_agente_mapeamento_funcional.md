# Agente 02 - Mapeamento Funcional

## Objetivo
Traduzir o comportamento da tela VB6 para linguagem funcional de negocio.

## Responsabilidades

- Descrever a finalidade real da tela
- Levantar o que o usuario pode fazer (acoes primarias e secundarias)
- Mapear fluxo funcional de ponta a ponta
- Identificar pre-condicoes, pos-condicoes e excecoes de uso
- Separar comportamento operacional de comportamento tecnico

## Entradas

- Saida do Agente 01
- Imagem da tela (se houver)
- Contexto operacional do modulo

## Saidas

- Documento funcional da tela
- Lista de acoes do usuario
- Fluxo funcional
- Regras de uso percebidas

## Checklist interno

- [ ] Finalidade da tela descrita em termos de negocio
- [ ] Acoes principais listadas
- [ ] Acoes auxiliares listadas
- [ ] Fluxo feliz descrito
- [ ] Fluxos alternativos/erro descritos
- [ ] Campos obrigatorios/opcionais inferidos
- [ ] Mensagens funcionais relevantes listadas

## Exemplo de uso

Entrada: diagnostico tecnico da "Associacao de Marcantes"

Saida esperada:
- Funcao: associar marcantes a cargas sem marcante
- Acoes: filtrar carga, incluir marcante, remover, salvar, sair
- Fluxo: selecionar carga -> montar lista temporaria -> consolidar associacao
