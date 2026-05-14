# Regras de negocio — Descarga DDC (Patio)

Legenda de classificacao:

- **RN**: Regra de negocio
- **RI**: Regra de interface
- **RT**: Regra tecnica
- **RV**: Regra de validacao
- **HP**: Hipotese (sem evidencia completa no codigo ou depende de confirmacao externa)

## Catalogo de regras

| ID | Descricao | Classificacao | Evidencia |
|----|-------------|---------------|-----------|
| R01 | Somente exibir placas vinculadas a ordens com `flag_saida=0` e `flag_ddc=1` | RN | `Carrega_Placa` — `FrmDescargaArmazem.frm` |
| R02 | Ao confirmar placa, deve existir gate com `flag_gate_in=1` e sem saida (`flag_gate_out` nulo/zero) alinhado por `placa` da ordem | RN + RV | `cbPlaca_LostFocus` |
| R03 | Lotes exibidos para selecao devem estar com `bl.flag_ativo=1` | RN | `cbPlaca_LostFocus` |
| R04 | "Carregar" exige placa e lote selecionados | RV | `CARREGAR` |
| R05 | Apos carregar, `TxtCntr` e `TxtOC` vêm do registro de saida (`tb_registro_saida_cntr`) | RN | `CARREGAR` |
| R06 | Itens da grade filtram `tb_descarga_armazem` por `cntr` e `gate` correntes | RN | `Carrega_Grid1` |
| R07 | Quantidade informada deve ser **> 0** | RV | `Salvar` |
| R08 | Embalagem obrigatoria | RV | `Salvar` |
| R09 | Deve indicar **SIM** ou **NAO** para entrega finalizada | RV | `Salvar` |
| R10 | `gate` e `lote` devem estar resolvidos (nao zero) para salvar | RV | `Salvar` |
| R11 | Em inclusao, grava usuario (`Cod_Usuario`) no item | RN | `Salvar` INSERT |
| R12 | Se `flag_ddc` do conteiner for 0, setar `flag_ddc=1` e `inicio_ddc=sysdate` | RN | `Salvar` pos-`Executa` |
| R13 | Se entrega finalizada **SIM** (`Option1`), setar `flag_ddc=2` no conteiner | RN | `Salvar` |
| R14 | Nao excluir item se existir `quantidade_saida` agregada em `tb_amr_gate` | RN + RV | `Excluir` |
| R15 | Nao permitir reduzir quantidade abaixo da soma ja saida em atualizacao | RV | `Salvar` UPDATE |
| R16 | Mensagem ao tentar trocar embalagem quando ha saida; fluxo SQL subsequente e ambiguo | RN + RT | `Salvar` UPDATE (bloco aninhado) |
| R17 | Atalhos F3/F5/F6 so disponiveis com `Frame1.Enabled` | RI | `Form_KeyDown` |
| R18 | Uso de `CONVERTE_SQL` e `Banco_Em_Uso` para ramificar Oracle vs outros | RT | `Carrega_Grid1`, `Form_Load`, INSERT |
| R19 | Textos de titulo inconsistentes entre menu, caption e `Label5` | RI | `.frm` |
| R20 | Impacto transacional: multiplas escritas em `Salvar` sem transacao declarada | RT | `Salvar` |
| R21 | Romaneio/carregamento pode ser afetado indiretamente por `tb_amr_gate` / ordens | HP | Nao ha referencia direta no `.frm` |
| R22 | Permissao especifica para o botao "Descarga DDC" | HP | Nao verificado neste `.frm` |

## Regras de consistencia e concorrencia

| Tema | Classificacao | Observacao |
|------|---------------|------------|
| Dependencia de `tb_gate_new` aberto | RN | Dois operadores concorrendo no mesmo veiculo — **HP**: pode haver corrida; validar no servidor com lock/versao. |
| Uso de recordset global `Rst` | RT | Risco de estado compartilhado entre telas no legado. |

## Regras implicitas inferidas (marcadas como hipotese)

| ID | Texto | Motivo |
|----|--------|--------|
| H01 | `flag_ddc` com valores 0/1/2 representa estados progressivos de DDC | Sem ENUM documentado no codigo |
| H02 | `inicio_ddc` nao e limpo ao alternar finalizacao | Nao ha UPDATE explicito para limpar |

## Inventario de mensagens (para homologacao)

Consolidar textos reais com usuarios (encoding ANSI corrompido no arquivo fonte).
