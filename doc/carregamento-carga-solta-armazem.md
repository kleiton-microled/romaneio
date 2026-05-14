# Documentacao de Uso - Carregamento Carga Solta Armazem (MVC)

## Objetivo da tela
Permitir ao operador de armazem realizar o carregamento de marcantes de carga solta em um veiculo selecionado, preservando as regras criticas da tela VB6 `CarregaCS.frm` (validacao de saida, controle de yard, historico de shifting e suporte a carregamento parcial).

## Quando utilizar
- Quando o veiculo ja agendado se apresenta no armazem para receber carga solta.
- Quando for necessario carregar (bipar) marcantes de uma Ordem de Carregamento (OC).
- Quando for necessario cancelar um carregamento ja realizado e devolver o marcante ao estoque do armazem.
- Quando o operador precisar consultar a localizacao fisica das cargas pendentes da OC.

## Como acessar
- Pelo menu principal: **Carregamento Carga Solta**.
- Ou diretamente pela rota: `~/CarregamentoCargaSolta/Index`.

## Pre-requisitos
- Estar autenticado no sistema (sessao valida).
- Estar logado no patio correto (a tela filtra os veiculos pelo `PATIO` do usuario; quando `PATIO = 1` ou `PATIO = 7`, a regra do legado expande para `PATIO IN (1, 7)`).
- Possuir Ordens de Carregamento criadas para o veiculo no patio.
- Quando o usuario tem `flag_carregamento_marcante_bip = 1` em `TB_CAD_USUARIOS`, a leitura do marcante so e aceita por bipagem (cole/digite manualmente bloqueado).

## Campos e componentes principais
- **Veiculo (placa cavalo)**: combo carregado com os veiculos do patio (`VW_CAM_CARREGAMENTO`). Mostra `placa cavalo + placa carreta + modelo`.
- **Atualizar veiculos**: recarrega o combo (apos um veiculo recem-agendado, por exemplo).
- **Ordens de Carregamento**: grid com as OCs do veiculo (numero, lote, item, qtde, qtde carregada, embalagem). Linhas verdes representam OCs ja totalmente carregadas.
- **Carregamentos da OC**: grid lateral com os marcantes ja carregados na OC selecionada e quantidade carregada.
- **Marcante**: campo principal de bipagem (9 digitos, somente numeros).
- **Lote / Qtde / Local / Armazem**: dados consultados automaticamente apos identificacao do marcante (`Lote` e `Armazem` somente leitura; `Qtde` editavel para suportar carregamento parcial).
- **Painel de critica**: mostra mensagens de regra de negocio (carga divergente, carga ja carregada, etc.).
- **Botoes**:
  - `[F1] Locais` - abre o modal com a posicao fisica das cargas da OC.
  - `[F2] Documentos` - aponta para os documentos da OC (LTL ou agendamento `TB_AG_CS`).
  - `[F3] Limpar` - limpa os campos do marcante e libera o input para nova bipagem.
  - `[F4] Carregar` - executa o carregamento do marcante na OC selecionada.
  - `[F5] Cancelar` - cancela o carregamento (devolve o marcante ao estoque).
  - `[ESC] Sair` - retorna ao menu inicial.

## Fluxo recomendado de operacao
1. Selecionar o veiculo na combo `Veiculo (placa cavalo)`.
2. Conferir as OCs listadas; selecionar a OC alvo (a coluna `Carregada` ajuda a identificar OCs em andamento).
3. Bipar (ou digitar) o marcante no campo `Marcante` e pressionar **Enter**.
4. Verificar:
   - se a `Qtde` precisa ser ajustada para carregamento parcial;
   - se o painel de critica esta limpo (sem mensagens vermelhas);
   - se `Lote` e `Armazem` correspondem ao esperado.
5. Acionar `[F4] Carregar`.
6. Conferir o feedback (`ITEM CARREGADO` ou `ITEM PARCIALMENTE CARREGADO`).
7. Repetir a partir do passo 3 ate concluir todas as cargas da OC.
8. Para cancelar um carregamento, bipar novamente o marcante (ja em `YARD = CAM`) e acionar `[F5] Cancelar`.

## Regras de validacao importantes
- **Veiculo obrigatorio**: a OC so e listada apos a selecao do veiculo.
- **Marcante existente**: marcante inexistente retorna `Marcante nao informado/invalido`.
- **OC compativel**: o `AUTONUM_CARGA` do marcante precisa ser igual ao `AUTONUMCS` da OC selecionada (mensagem `Carga Divergente da Ordem de carregamento`).
- **Carga ja carregada**: marcantes em `YARD = CAM` so aceitam **Cancelar**, nao **Carregar** (mensagem `Carga ja carregada`).
- **Quantidade**: `Qtde` deve ser maior que 0 e menor ou igual ao `VOLUMES` atual do marcante (caso contrario `Quantidade indisponivel`).
- **Funcao de saida**: antes de `Carregar`, o backend executa `dbo.FC_VALIDA_SAIDA_CARGA(BL, ORDEM, AUTONUMCS)`. Qualquer retorno diferente de `OK` e exibido como critica e o carregamento e bloqueado.
- **Bipagem obrigatoria**: quando `Wflag_carregamento_marcante_bip = 1`, colar conteudo no campo `Marcante` e bloqueado e o input fica `readonly` apos a validacao (so e liberado por `[F3] Limpar`).
- **Patio**: a tela so opera com veiculos do patio do usuario (regra `PATIO` do legado).

## Mensagens comuns
- **Validacao de entrada**:
  - `Marcante nao informado/invalido`
  - `Quantidade indisponivel`
  - `Veiculo nao selecionado`
  - `Selecione uma ordem de carregamento`
- **Restricao operacional**:
  - `Carga Divergente da Ordem de carregamento`
  - `Carga ja carregada` / `Carga ainda nao carregada`
  - `Lote nao encontrado`
  - Mensagens devolvidas por `dbo.FC_VALIDA_SAIDA_CARGA` (variam por contexto)
- **Resultado**:
  - `ITEM CARREGADO`
  - `ITEM PARCIALMENTE CARREGADO`
  - `ITEM CANCELADO`
- **Documentos**:
  - `Documento LTL para o lote N` (quando ha solicitacao LTL)
  - `Agendamento N localizado para o lote N` (caso geral)
  - `Nenhum agendamento de documento encontrado para o lote informado`

## Atalhos de teclado
- `F1` Locais
- `F2` Documentos
- `F3` Limpar
- `F4` Carregar
- `F5` Cancelar
- `ESC` Sair
- `Enter` no campo `Marcante` confirma a leitura e dispara a busca dos dados.

> Observacao: no ambiente web, atalhos podem variar por navegador/politica local. Sempre use os botoes como fallback.

## Botao Documentos (F2)
- O endpoint `CarregamentoCargaSolta/Documento` identifica se o lote da OC tem solicitacao LTL ativa (`TB_SOLICITACAO_LTL.flag_ltl = 1`) ou retorna o `MAX(AUTONUM)` de `TB_AG_CS` para o lote, e devolve `idSolicitacaoLtl` quando for LTL.
- No legado VB6 (`ProjetoReferencia/Coletor/CarregaCS.frm`, `cmdF2_Click`), apos essas consultas o sistema chama `Shell` com `ConsultaDocumentos\ConsultaDocColetor.exe` (projeto fonte em `Band .NET\ConsultaDocColetor`).
- No **Romaneio MVC**, o botao `[F2] Documentos` abre em nova aba o modulo compartilhado **`ConsultaDocumentosAgendamento/Index`** (`doc/consulta-documentos-agendamento.md`), que lista documentos e exibe PDF/imagem quando o blob existe em `TB_AG_DOCUMENTOS`.
- Arquivos obtidos apenas via servico **Bandeirantes.AnexoDiretorio** (sem blob na tabela) ainda nao sao exibidos no web ate integrar esse servico.

## Boas praticas de uso
- Sempre verificar a coluna `Carregada` na grid de OCs antes de iniciar a bipagem.
- Em carregamento parcial, ajustar a `Qtde` antes de `[F4] Carregar`. O sistema atualiza `VOLUMES` do marcante e a `tb_temp_marcante` automaticamente.
- Use o `[F1] Locais` para orientar o conferente fisico antes de buscar os volumes.
- Em caso de critica vermelha, leia a mensagem antes de tentar novamente; ela costuma indicar diretamente a regra de negocio violada.
- Para sair da tela, prefira `[ESC] Sair` para garantir o retorno limpo ao menu.

## Resultado esperado
Com OC valida, marcante valido, quantidade compativel e funcao `FC_VALIDA_SAIDA_CARGA` retornando `OK`, o carregamento e gravado em transacao unica:

- `TB_CARGA_SOLTA_YARD` recebe um novo registro com `YARD = CAM` (quantidade carregada).
- `TB_MARCANTES` e religado ao novo registro yard e recebe `PLACA_C` e `ORDEM`.
- Em carregamento parcial, `VOLUMES` em `TB_MARCANTES` e a quantidade em `tb_temp_marcante` sao atualizadas.
- `TB_HIST_SHIFTING_CS` recebe o evento de shifting (`YARD = CAM`).

No cancelamento, `TB_CARGA_SOLTA_YARD` ganha um registro `YARD = CANCC`, o marcante volta ao estoque (com `PLACA_C = NULL`) e o evento e historiado em `TB_HIST_SHIFTING_CS` (`YARD = CANCC`).

---

## Pendencias funcionais (validar antes do go-live)

Estas pendencias foram identificadas durante a migracao (Agente 07) e estao listadas tambem em `MigracaoVB6/Outputs/carregamento-carga-solta-armazem/00_consolidado.md`:

1. **MOTIVO = 8 fixo nos inserts de `TB_CARGA_SOLTA_YARD`**
   - O legado utiliza `MOTIVO = 8` tanto no carregamento quanto no cancelamento.
   - Confirmar com a area de negocio se `8` corresponde realmente a "Carregamento" e se faz sentido reaproveitar o mesmo motivo no `CANCC`.

2. **Historico em `[F5] Cancelar`**
   - Adicionado por decisao tecnica registro em `TB_HIST_SHIFTING_CS` com `YARD = CANCC` (rastreabilidade).
   - O legado VB6 nao gravava esse evento. Validar se o valor `CANCC` e aceito pelos relatorios e pelas demais telas que consomem o historico.

3. **Carregamento parcial e `tb_temp_marcante`**
   - O `UPDATE QTDE` em `tb_temp_marcante` ocorre nos casos parciais para manter compatibilidade com a tela `Associacao de Marcantes`.
   - Validar se essa interacao continua coerente com o fluxo da `Associacao de Marcantes` no MVC.

4. **Permissionamento equivalente a `Valida_Acesso_Botao` (perfil `COLETOR_ARM`)**
   - O legado consulta `TB_SYS_FUNCOES`, `TB_SYS_GRP_PERMISSOES` e `TB_SYS_USER_GRUPOS` para autorizar o acesso a tela.
   - O MVC atualmente protege apenas pela sessao (`Session["Logado"]`). Definir se a matriz de permissao deve ser aplicada e onde (filter MVC, decorator no controller etc.).

5. **`[F2] Documentos` em ambiente web**
   - No `ProjetoReferencia`, o `[F2] DOC` em `CarregaCS.frm` dispara `ConsultaDocumentos\ConsultaDocColetor.exe` (fonte em `Band .NET\ConsultaDocColetor`).
   - **Implementado:** modulo `ConsultaDocumentosAgendamento` no Romaneio (ver `doc/consulta-documentos-agendamento.md`). Pendente: integrar **Bandeirantes.AnexoDiretorio** quando `DOCUMENTO` na tabela for nulo (comportamento do WinForms).

6. **Disponibilidade de `dbo.FC_VALIDA_SAIDA_CARGA`**
   - A funcao escalar e chamada em todo `[F4] Carregar`. Confirmar a presenca dessa function no banco do ambiente alvo (homologacao e producao). Em caso de ausencia, o carregamento ira falhar com erro de SQL.

7. **`Wflag_carregamento_marcante_bip` por usuario**
   - A flag e lida em `Login` e gravada em `Session["FlagCarregamentoMarcanteBip"]`. Validar se o cadastro de usuarios atualizado contempla os operadores de carregamento corretos com a flag ativa quando exigido.

## Pendencias tecnicas

1. **Build local**
   - O ambiente de desenvolvimento utilizado para a migracao nao tinha Visual Studio/MSBuild instalado, portanto a checagem foi feita estaticamente (lints e revisao manual). Recomenda-se:
     - abrir `Romaneio.sln` no Visual Studio 2019/2022;
     - confirmar que os novos arquivos foram detectados como Compile/Content;
     - executar um build da configuracao Debug.

2. **Modal de Locais**
   - A view inclui `~/Content/js/bootstrap.min.js` localmente para suportar `$.modal('show')`. Caso o tema/layout do `Romaneio` evolua para Bootstrap 4/5, este script precisara ser revisto.

3. **Mensagens em portugues sem acentos**
   - Para alinhamento com o padrao das demais telas migradas, as mensagens foram inseridas sem acentos. Caso a aplicacao adote internacionalizacao no futuro, esses textos podem ser migrados para resources.

4. **Testes automatizados**
   - O `Romaneio` nao possui projeto de testes. Quando criado, sugere-se cobrir os cenarios:
     - listagem de veiculos por patio (1/7);
     - listagem de OCs por placa com carregamento parcial;
     - `Carregar` com `FC_VALIDA_SAIDA_CARGA` retornando `OK` / != `OK`;
     - `Carregar` com `Carga Divergente`;
     - `Cancelar` revertendo o yard e gravando `CANCC`.

5. **Logs/observabilidade**
   - O fluxo nao registra logs detalhados das operacoes (apenas o historico de shifting na tabela). Avaliar a inclusao em uma camada de logging estruturada quando o `Romaneio` adotar essa pratica.
