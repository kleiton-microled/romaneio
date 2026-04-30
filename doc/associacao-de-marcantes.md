# Documentacao de Uso - Associacao de Marcantes (MVC)

## Objetivo da tela
Permitir a associacao de marcantes a cargas pendentes, seguindo o mesmo fluxo funcional da tela legada em VB6.

## Quando utilizar
- Quando houver cargas sem marcante vinculado.
- Quando for necessario incluir, remover ou confirmar marcantes antes da gravacao final.
- Quando a operacao exigir tratamento de carga em modo exportacao (EXP), conforme retorno do filtro.

## Campos e componentes principais
- **Lote**: campo inicial para localizar as cargas.
- **CNTR**: filtro complementar por conteiner.
- **Grid de cargas**: lista os itens retornados pelo filtro.
- **Indicador EXP**: exibido quando a origem da consulta estiver no contexto de exportacao.
- **Marcante**: codigo do marcante a ser associado.
- **Qtde**: quantidade vinculada ao marcante informado.
- **Lista temporaria de marcantes**: staging com os marcantes adicionados antes de salvar.

## Fluxo recomendado de operacao
1. Informar o **Lote**.
2. (Opcional) Selecionar o **CNTR**.
3. Acionar **Filtrar** para carregar o grid de cargas.
4. Selecionar uma linha no grid.
5. Informar **Marcante** e **Qtde**.
6. Acionar **Adicionar** para enviar para a lista temporaria.
7. Repetir a inclusao para todos os marcantes necessarios.
8. Revisar a lista temporaria.
9. Acionar **Salvar** para persistir as associacoes.

## Atalhos de teclado
- **F1**: adiciona marcante na lista temporaria.
- **F2**: remove item selecionado da lista temporaria.
- **F4**: executa filtro.
- **F5**: salva as associacoes.
- **ESC**: sai da tela.

> Observacao: no ambiente web, atalhos podem variar conforme navegador/politicas locais. Sempre existe fallback pelos botoes da tela.

## Validacoes esperadas
- Marcante obrigatorio para inclusao.
- Quantidade obrigatoria e valida.
- Bloqueio de duplicidade de marcante na lista temporaria (quando aplicavel a regra).
- Consistencia de quantidade validada antes da gravacao final.
- Regras especificas de exportacao aplicadas quando o modo **EXP** estiver ativo.

## Boas praticas de uso
- Sempre selecionar a carga correta antes de adicionar marcantes.
- Conferir a lista temporaria antes de salvar.
- Em caso de erro de validacao, corrigir os dados e repetir a operacao.
- Evitar uso exclusivo de atalho; preferir botoes quando houver conflito com o navegador.

## Resultado esperado
Ao salvar com sucesso, os marcantes da lista temporaria sao persistidos para a carga selecionada, mantendo integridade das regras de negocio.
