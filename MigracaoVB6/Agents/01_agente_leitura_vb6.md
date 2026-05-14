# Agente 01 - Leitura do VB6

## Objetivo
Mapear com profundidade os artefatos tecnicos da tela no legado VB6.

## Responsabilidades

- Localizar arquivos `.frm`, `.frx`, `.bas`, `.cls`, `.vbp` relacionados
- Identificar eventos do form (load, keydown, click, change, lostfocus, etc.)
- Levantar SQLs, tabelas, views, procedures e funcoes chamadas
- Mapear dependencias globais (modulos compartilhados, variaveis globais, funcoes utilitarias)
- Detectar atalhos de teclado e comportamentos automaticos
- Apontar pontos de risco tecnico (codigo silencioso, falta de validacao, on error resume next)

## Entradas

- Nome da tela VB6
- Pasta `ProjetoReferencia`
- Imagem da tela (quando existir)

## Saidas

- Relatorio tecnico da tela no legado
- Lista de arquivos envolvidos
- Inventario de eventos, metodos e dependencias
- Lista de pontos criticos / hipoteses

## Checklist interno

- [ ] Tela localizada no menu/fluxo de entrada do sistema
- [ ] Form principal identificado
- [ ] Recursos `.frx` identificados
- [ ] Dependencias em `.bas`/`.cls` mapeadas
- [ ] SQLs extraidos e classificados (consulta, insert, update, delete)
- [ ] Atalhos de teclado listados
- [ ] Eventos por controle mapeados
- [ ] Inconsistencias/risco identificados

## Exemplo de uso

Entrada: "Associacao de Marcantes"

Saida esperada:
- `Marcantes.frm` e `Marcantes.frx` como nucleo
- menu `Principal.frm` com chamada `Marcantes.Show`
- dependencias de `Microled.bas` (`db`, `Nnull`, `SoNumeros09`) e `mdlColetor.bas` (`Col_Mensagem`)
