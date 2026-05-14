# Consolidado - Localizacao Carregamentos

## Resumo executivo
- Tela: Localizacao Carregamentos (`Localizacao.frm` / `LocalizacaoS.frm`).
- Objetivo: localizar carregamentos ativos no patio por placa/carreta e exibir dados operacionais de posicao/liberacao.
- Status do estudo: concluido (documentacao completa), sem implementacao MVC.

## Principais achados
- Achado 1:
  - Existem duas variantes da tela com diferencas reais de comportamento e SQL (`Localizacao` e `LocalizacaoS`).
- Achado 2:
  - Consulta central usa `UNION ALL` para juntar fluxo de conteiner e carga solta, com totalizadores por tipo.
- Achado 3:
  - Na variante `S`, existe acao de movimentacao para `SAIDA` com validacoes de empilhamento e historico.

## Lista de arquivos VB6 relacionados
- `ProjetoReferencia/Coletor/LocalizacaoS.frm`
- `ProjetoReferencia/Coletor/LocalizacaoS.frx`
- `ProjetoReferencia/Coletor/Localizacao.frm`
- `ProjetoReferencia/Coletor/Localizacao.frx`
- `ProjetoReferencia/Coletor/PrincipalS.frm`
- `ProjetoReferencia/Coletor/PrincipalG.frm`
- `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp`
- `ProjetoReferencia/Coletor/ColetorPatioNaoUsando.vbp`
- `ProjetoReferencia/Coletor/Microled.bas` (dependencias globais por uso)
- `ProjetoReferencia/Coletor/mdlColetor.bas` (dependencias globais por uso)

## Regras criticas
- Regra:
  - Somente registros sem `dt_gate_out` e dentro da janela temporal devem ser exibidos.
- Regra:
  - Filtro por placa deve aplicar em `placa` e `carreta`.
- Regra:
  - `SAIDA` so pode ocorrer para conteiner, nao para carga solta.
- Regra:
  - Se ja estiver em `SAIDA` ou houver conteiner acima na pilha, bloquear movimentacao.

## Objetos de banco envolvidos
- Tabelas:
  - `operador..tb_gate_new`, `operador..tb_amr_gate`, `operador..TB_CAD_TRANSPORTADORAS`, `operador..tb_remocao`, `operador..tb_patios`, `operador..TB_YARD`
  - `sgipa..TB_CNTR_BL`, `sgipa..TB_CARGA_SOLTA`, `sgipa..TB_ORDEM_CARREGAMENTO`, `sgipa..tb_registro_saida_cs`, `sgipa..tb_cad_usuarios`, `sgipa..TB_HIST_SHIFTING`
- Views:
  - `sgipa..VW_LOCAL_OC`, `operador..VW_INVENT_SISTEMAS`
- Campos gravados:
  - `TB_CNTR_BL.YARD`
  - `TB_HIST_SHIFTING` (registro de movimentacao)
- Joins relevantes:
  - Gate -> AMR -> Ordem -> Carga/Conteiner + dados de transportadora/usuario.
- Filtros obrigatorios:
  - `dt_gate_out is null`, janela temporal, patio da sessao.
- Concorrencia:
  - no legado nao ha transacao explicita para update+insert de `SAIDA`.

## Decisoes de migracao
- Decisao:
  - Criar modulo MVC dedicado (`LocalizacaoCarregamentosController` + repositorio + view + js) seguindo padrao atual de `Romaneio`.
- Justificativa:
  - Consulta e regras da tela sao especificas e com impacto operacional.

- Decisao:
  - Preservar comportamento VB6 com SQL parametrizado e retorno JSON padrao.
- Justificativa:
  - Reduz risco tecnico sem perder regra de negocio.

## Criterios de aceite da migracao
- Filtros:
  - Filtrar por placa/carreta e patio deve retornar os mesmos registros do VB6.
- Campos exibidos:
  - cavalo, carreta, lib. ent., gate in, carga, quadra, rua, transportadora, responsavel.
- Acoes:
  - filtrar, selecionar linha, sair com `ESC`; e `SAIDA` se esta regra permanecer no escopo final.
- Validacoes:
  - bloquear `SAIDA` para item ja em SAIDA.
  - bloquear `SAIDA` com conteiner acima.
- Regras preservadas/adaptadas:
  - manter regra por patio e totalizadores CNTR/CS.
  - adaptar mensagens/confirmacao para modal web sem perder semantica.

## Pendencias de validacao
- Pendencia:
  - Definir variante oficial para baseline (`Localizacao.frm` vs `LocalizacaoS.frm`).
- Responsavel:
  - negocio + time tecnico + DBA.

- Pendencia:
  - Confirmar janela temporal oficial (1 dia ou 3 dias).
- Responsavel:
  - operacao/negocio.

## Riscos e lacunas
- Risco:
  - Divergencia entre variantes pode gerar implementacao incorreta da consulta.
- Risco:
  - Falta de transacao no legado mascara inconsistencias que podem surgir em falha parcial.
- Lacuna:
  - Nao houve evidencia completa do codigo-fonte das globais em modulos compartilhados externos.
- Lacuna:
  - Sem referencia de imagem oficial da tela nesta etapa; mapeamento visual foi feito do `.frm`.

## Escopo da implementacao (fase atual)
- Incluido:
  - diagnostico, mapeamento funcional, mapeamento visual, regras, desenho MVC e plano.
- Fora de escopo:
  - implementacao MVC, testes de codigo e deploy.

## Proximas etapas
1. Validar baseline funcional com usuario-chave (variante e janela temporal).
2. Rodar prompt do Agente 07 para implementacao MVC.
3. Homologar equivalencia funcional contra roteiro de aceite.

## Prontidao para implementacao
- Nivel de prontidao: Medio-Alto.
- Condicoes para iniciar:
  - fechar pendencias de baseline (variante e janela temporal).
  - confirmar escopo final da acao `SAIDA` na versao web.
