# 02 — Mapeamento funcional — Movimentação Container: Pátio

## Finalidade da tela (negócio)

Permitir que o operador do **pátio**:

1. **Localize** um contêiner no estoque do pátio informando o **final da sigla** ou o **ID completo**.
2. **Consulte** dados operacionais do equipamento (navio/viagem, yard atual, tipo/tamanho, lacres, IMO, Reefer, indicadores de bloqueio, SIGVIG, agendamentos, vínculos com BLs).
3. **Planeje / confirme movimentação de yard** informando **destino**, **motivo** e, quando exigido, **câmera** de posicionamento.
4. **Execute** a movimentação persistindo via **`pcMovimentaCntr`**, após validação **`fnMovimentaCntr`**.
5. Tratar fluxos especiais (**SAÍDA**, **DTA**, **BK**, **CLIA**, segregação **IMO**, divergência de lacres, bloqueios Reefer/scanner).
6. Na variante **Pátio Small**, **registrar e finalizar avarias** do contêiner em aba dedicada, sujeito a **permissões** no cadastro de funções.

A tela **não é só consulta**: a ação principal **altera posição / status operacional** do contêiner via stored procedure e efeitos colaterais (histórico, avarias automáticas, possível estorno).

---

## Atores e pré-condições

| Ator | Pré-condição |
|------|----------------|
| Operador autenticado no Coletor | `Cod_Usuario`, `Col_Patio` válidos na sessão global |
| Permissões de avarias (opcional) | Funções **998** (acesso/excluir/incluir) e **999** (finalizar), conforme queries no `Form_Load` |

---

## Fluxo principal (feliz)

1. Operador abre a tela pelo menu **Movimentação Conteiner**.
2. Campo **final** (`CNTR`): sistema resolve 0/1/N candidatos via `fnInvent_Sistemas_YardF`.
3. Com ID completo em `MskCntr`, no `LostFocus` carrega **Busca_Dados**:
   - Preenche painel de dados e mensagens contextuais (`Text1(0)`).
   - Carrega grids de próximos movimentos (`Carrega_Grid1`) e vizinhos de pilha (`Carrega_Grid2`).
   - Amarra `DataGrid2` aos BLs do contêiner.
4. Operador informa **motivo** (`DataCombo2`) e **destino** (`mskDestino`).
5. Clica **Salvar** (`Command1`):
   - Se motivo exige **câmera**, abre frame de seleção e aguarda confirmação (`Command3`).
   - Caso contrário chama **`Atualiza_Posicao`**.
6. `Atualiza_Posicao` valida regras (lacres, Reefer, motivo, destino ≠ origem, bloqueios, BK/CLIA, scanner na saída, etc.).
7. Chama **`fnMovimentaCntr`**; se retorno ≠ OK, exibe mensagem e aborta.
8. Se OK, executa **`pcMovimentaCntr`** com lacros e parâmetros coletados.
9. Trata divergência de lacres (pode inserir registro em **`TB_AVARIAS_CONTEINER`**).
10. Verifica segregação IMO; se violação, **estorna** via segunda execução da SP com yards invertidos.
11. Limpa formulário e reposiciona foco para novo contêiner.

---

## Fluxos alternativos e erro

| Situação | Comportamento evidenciado |
|----------|---------------------------|
| Nenhum ou vários contêineres para o mesmo final | Mensagens modais; reinicia máscara |
| Contêiner não encontrado em `fnInvent_Sistemas_Col` | “Unidade não encontrada no estoque” |
| Destino vazio ou igual ao yard atual | Bloqueio com mensagem |
| Saída (`mskDestino = "SAIDA"`) | Confirmação Yes/No; checagem `FN_BLOQUEIO_SCANNER` para sistema `I` |
| Reefer sem temperatura ligado | Bloqueio de posicionamento |
| Lacres obrigatórios não informados | Mensagens específicas (inclui fluxo `I` com checagem de histórico shifting ≠ GATEIN e checkbox **LACRES CONFERIDOS**) |
| Motivo com flag de câmera | Interrompe para escolher `DataCombo4` |
| Destinos BK / CLIA | Valida flags em `TB_CNTR_BL` e janelas de reserva/agenda; pode atualizar `AUTONUM_GD_RESERVA` |
| EF=E e SAÍDA | Abre **`FrmVeiculoServ`** antes da movimentação |
| Destino DTA | Abre **`FrmVeiculoServ`**; após uso força destino textual **CAM** |
| Avarias finalizadas | Botões de inclusão/exclusão/finalização desabilitados ou bloqueados por mensagem |

---

## Dados consultados (somente leitura na tela)

- Inventário por pátio: TVFs `fnInvent_Sistemas_YardF`, `fnInvent_Sistemas_Col`
- Próximos movimentos / fila: views `VW_PROX_MOV_*`
- Cadastros: motivos, câmeras, yards válidos, descritivo do pátio
- BLs associados ao contêiner (grid)
- Histórico de shifting (aba Histórico)
- Histórico e lista de avarias (aba Avarias + `VW_EGMCI_AVARIAS_TTL`)
- Regras auxiliares: bloqueios alfândega/manual, registro de carregamento, SIGVIG, parceiro transporte interno (`Frame5` condicional)

---

## Dados alterados (persistência)

- **Movimentação principal:** via **`pcMovimentaCntr`** (efeitos internos da SP — não desmontados neste estudo)
- **Atualização pontual:** `TB_CNTR_BL.AUTONUM_GD_RESERVA` em cenário BK (trecho específico)
- **Avarias:** `INSERT`/`DELETE` em `TB_AVARIAS_CONTEINER`; `INSERT` em `TB_ETAPAS_AVARIA` na finalização
- **Avaria automática:** divergência lacres × manifesto

---

## Campos de entrada vs somente leitura (visão funcional)

| Grupo | Entrada pelo usuário | Somente leitura / sistema |
|-------|----------------------|---------------------------|
| Identificação | Final (`CNTR`), ID completo (`MskCntr`) | `txtPatio`, vários campos de carga após busca |
| Movimentação | `mskDestino`, `DataCombo2` (motivo), `DataCombo4` (câmera), `Check1` | `TXTYardAtual` |
| Lacres / peso | Lacres IPA/Gate/Lona, tara (conforme fluxo) | ISO, max gross em alguns contextos |
| Apoio | Opções ATUAL/DESTINO para grid 2 | Grids e labels de status |

---

## Mensagens funcionais (exemplos evidenciados)

- “Nenhum contêiner com este final” / “Existe mais de 1 contêiner…”
- “Unidade não encontrada no estoque”
- “Informe o motivo” / “Informe o motivo da movimentação!”
- “Destino não Informado!” / “Destino e Origem iguais!”
- “Confirma Saida da Unidade ?”
- “Esta unidade tem BLOQUEIO de ESCANER”
- “Contêiner Reefer sem temperatura - Posicionamento bloqueado”
- “Pelo menos 1 lacre de origem deve ser informado”
- “Informe que os lacres foram conferidos”
- “Lacres informados incorretamente”
- “SEGREGACAO IMO !!!! …” (após estorno)
- Avarias: “Nenhum Local selecionado”, “Informe apenas um local”, “Já consta finalização do Patio para a unidade”, etc.

---

## Impacto em processos adjacentes

| Processo | Impacto |
|----------|---------|
| **Pátio / yard** | Atualização explícita de posição e histórico de shifting |
| **Contêiner** | Validações de Reefer, scanner, BK/CLIA, DTA, lacres, IMO |
| **Romaneio / carregamento** | Consultas a ordem de carregamento / registro de saída para **alerta** ao operador (não conclusivo de gravação direta na tela além da SP) |

---

## Hipóteses

1. **Toda** persistência de estado de yard passa por `pcMovimentaCntr` — **hipótese forte**, não inspecionamos o corpo da procedure no repositório atual.
2. O modal `FrmVeiculoServ` pode gravar dados adicionais (veículo/serviço) necessários para SAÍDA/DTA — **hipótese**; requer leitura desse `.frm` numa iteracao futura se for obrigatório para paridade.
