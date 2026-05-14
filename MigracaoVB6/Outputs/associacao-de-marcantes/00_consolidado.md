# Template 00 - Consolidado da Tela

## Resumo executivo
- Tela: Associacao de Marcantes
- Objetivo: associar marcantes a cargas sem marcacao, com controle de quantidade e contexto normal/EXP.
- Status do estudo: concluido (fase documental), sem implementacao MVC.

## Principais achados
- Achado 1:
  - a tela usa duas fontes de dados (`VW_WMS_SEM_MARCANTE` e fallback `VW_WMS_SEM_MARCANTE_CEXP`) e muda regras no modo EXP.
- Achado 2:
  - a operacao eh baseada em staging via `TB_TEMP_MARCANTE`/`TB_TEMP_MARCANTE_CEXP` antes da consolidacao final.
- Achado 3:
  - existem riscos tecnicos no legado (SQL dinamico, erros silenciosos, validacao de selecao fragil) que devem ser corrigidos na migracao.

## Regras criticas
- Regra:
  - soma das quantidades dos marcantes em staging deve ser igual a quantidade do item selecionado.
- Regra:
  - marcante de exportacao so pode ser associado em modo EXP e com `AUTONUM_CEXP` compatvel.

## Decisoes de migracao
- Decisao:
  - criar modulo MVC dedicado (`AssociacaoMarcantesController` + repositorio proprio).
- Justificativa:
  - regras e staging especificos nao encaixam bem como extensao simples de outro modulo.

- Decisao:
  - manter staging em `TB_TEMP_*` na primeira versao.
- Justificativa:
  - maior fidelidade comportamental com o legado na fase inicial.

## Pendencias de validacao
- Pendencia:
  - confirmar politica de limpeza da `TB_TEMP_MARCANTE` e `TB_TEMP_MARCANTE_CEXP`.
- Responsavel:
  - time funcional/DBA (nao evidenciado no form VB6).

- Pendencia:
  - validar origem das setas laterais vistas na imagem (nao presentes no `.frm`).
- Responsavel:
  - time funcional com acesso ao executavel legado.

## Escopo da implementacao (fase atual)
- Incluido:
  - diagnostico tecnico completo, mapeamento funcional/visual, regras classificadas, desenho MVC e plano em etapas.
- Fora de escopo:
  - codificacao de controller/view/repositorio, alteracoes de banco, testes automatizados de implementacao.

## Riscos e lacunas de entendimento
- Risco:
  - comportamento parcial no salvar quando condicao de capacidade nao permite update de um item, sem mensagem explicita.
- Risco:
  - divergir da operacao legacy nos atalhos F-key se nao houver controle adequado no browser.
- Lacuna:
  - ausencia de evidencias sobre limpeza global de tabelas temporarias.

## Proximas etapas
1. Validar pendencias abertas com operacao/DBA.
2. Iniciar implementacao pela Etapa 1 do plano (`06_plano_implementacao.md`).
3. Homologar paridade funcional com usuarios da operacao apos primeira entrega MVC.

## Prontidao para implementacao
- Nivel de prontidao: Alto (documentacao suficiente para comecar desenvolvimento com baixa ambiguidade).
- Condicoes para inicio seguro:
  - confirmar pendencias de temp e setas visuais.
  - alinhar estrategia de atalhos web com equipe funcional.
