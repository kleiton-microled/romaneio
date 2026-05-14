# Plano de implementacao — Descarga DDC (Patio)

## Principios

- Entregar em **fatias** que permitam homologar consultas antes de liberar gravacao.
- **Parametrizar SQL** e introduzir **transacao** no fluxo equivalente a `Salvar`.
- Nao divergir do padrao de DI/controllers do Romaneio.

## Fases sugeridas

### Fase 0 — Pre-requisitos de dados e ambiente

| Item | Saida |
|------|--------|
| Confirmar dialeto SQL do Romaneio para tabelas `SGIPA`/`OPERADOR` | Decisao documentada |
| Mapear chaves e tipos de `tb_descarga_armazem`, `tb_cntr_bl.flag_ddc` | ERD resumido |
| Validar semantica de `flag_ddc` (0/1/2) com negocio | Documento curto |

**Criterio de pronto**: equipe alinha modelo de dados e nomenclatura publica da tela.

### Fase 1 — Repositorio somente leitura

| Entrega | Detalhe |
|---------|---------|
| Interface + repositorio | `IDescargaDdcPatioRepositorio` |
| Consultas | Placas DDC, lotes ativos, lista de itens, embalagens |
| Testes manuais | Comparar resultados com telas VB6 em ambiente de homologacao |

**Criterio de pronto**: endpoints retornam os mesmos conjuntos que o legado para casos de teste escolhidos.

### Fase 2 — Controller + View esqueleto

| Entrega | Detalhe |
|---------|---------|
| `DescargaDdcPatioController.Index` | Auth + modelo basico |
| View Razor | Layout cabecalho + area de itens desabilitada ate carregar |
| JS | Chamadas AJAX para fase 1 |

**Criterio de pronto**: navegacao e carregamento de dados sem gravacao.

### Fase 3 — Operacoes de escrita

| Entrega | Detalhe |
|---------|---------|
| `SalvarItem` | INSERT/UPDATE com validacoes do legado |
| Atualizacao `tb_cntr_bl` | Mesma ordem logica (inicio DDC + finalizado) |
| `ExcluirItem` | DELETE com bloqueio por saida |
| Transacao | Envolver operacoes relacionadas de um mesmo submit |

**Criterio de pronto**: paridade funcional com VB6 em casos felizes e de erro esperados.

### Fase 4 — UX, acessibilidade e hardening

| Entrega | Detalhe |
|---------|---------|
| Atalhos | Opcional; documentar conflitos |
| Mensagens | Textos corrigidos (encoding) e i18n se aplicavel |
| Seguranca | Checagens de patio/conta do usuario se exigido |

**Criterio de pronto**: checklist de aceite do `00_consolidado.md` verde.

## Dependencias entre fases

```mermaid
graph LR
  F0[Fase 0] --> F1[Fase 1]
  F1 --> F2[Fase 2]
  F2 --> F3[Fase 3]
  F3 --> F4[Fase 4]
```

## Checklist tecnico (execucao)

- [ ] SQL parametrizado em todos os metodos do repositorio
- [ ] Transacao no salvar (item + `tb_cntr_bl`)
- [ ] Tratamento de concorrencia (versao ou re-leitura antes de update) — definir na Fase 0
- [ ] Logs de auditoria — **Hipotese**: padrao Romaneio pode exigir tabela de log
- [ ] Registro Unity em `Global.asax.cs`
- [ ] Link no `Home.cshtml`
- [ ] Testes de regressao comparando quantidades com `tb_amr_gate`

## Plano de testes (alto nivel)

1. Placa sem gate aberto — deve bloquear com mensagem equivalente.
2. Carregar com lote inativo — nao deve aparecer na lista (indiretamente validado pelo SQL).
3. Inserir item — linha aparece na lista; `flag_ddc` vai para 1 quando era 0.
4. Marcar finalizado SIM — `flag_ddc=2`.
5. Tentar excluir com saida — bloquear.
6. Tentar reduzir quantidade abaixo da saida — bloquear.

## Riscos por fase

| Fase | Risco |
|------|-------|
| F0 | Semantica incorreta de `flag_ddc` leva a status errados |
| F1 | Divergencia Oracle vs SQL Server nos filtros de data/null |
| F3 | Inconsistencia se nao usar transacao |
| F4 | Atalhos conflitando com navegador |
