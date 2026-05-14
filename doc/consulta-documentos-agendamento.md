# Documentacao de Uso - Consulta Documentos Agendamento (MVC)

## Origem
Modulo web equivalente ao projeto WinForms **ConsultaDocColetor** (`ConsultaDocumentos.exe` / pasta `ConsultaDocumentos` citada pelo coletor VB6 em `CarregaCS.frm`, `cmdF2_Click`).

## Objetivo
Listar documentos de agendamento (`TB_AG_DOCUMENTOS` + `TB_AGENDAMENTO_DOC`) e permitir visualizacao quando o binario existir em `TB_AG_DOCUMENTOS.DOCUMENTO`.

## Quando utilizar
- Tela **Carregamento Carga Solta Armazem**: botao `[F2] Documentos` (abre este modulo em nova aba).
- Qualquer outra tela MVC que possua `LOTE` (BL) e, no fluxo carga solta, o `AUTONUM` de `TB_AG_CS`, ou fluxo LTL com `LOTE` (e opcionalmente `ID` da `TB_SOLICITACAO_LTL`).

## Como abrir (URL)
Base: `~/ConsultaDocumentosAgendamento/Index`

### Carga solta (nao LTL)
```
/ConsultaDocumentosAgendamento/Index?ltl=0&lote={AUTONUM_BL}&ag={AUTONUM_TB_AG_CS}
```

### LTL
```
/ConsultaDocumentosAgendamento/Index?ltl=1&lote={AUTONUM_BL}
```
Opcional: `&idLtl={ID}` da `TB_SOLICITACAO_LTL` se ja conhecido (evita segunda consulta no servidor).

## Comportamento
- **Lista**: replica a consulta unificada do formulario `FrmDocumentosAgendamento.Consultar` (portada para SQL Server: `CONVERT` no lugar de `TO_CHAR`).
- **Bloqueio** (fluxo carga solta, alinhado ao `Program.vb` WinForms): ao abrir a tela com `ltl=0` e `ag` valido, insere em `OPERADOR..TB_ADMPATIO_BLOQUEIO_DOC`. O botao **Fechar e liberar bloqueio** ou **Voltar ao sistema** remove o registro para o mesmo usuario da sessao.
- **Visualizacao**: acao `Arquivo` le `DOCUMENTO` + `EXTENSAO` em `SGIPA..TB_AG_DOCUMENTOS`. Se `DOCUMENTO` for nulo, exibe mensagem explicando que no legado o arquivo vinha do servico **Bandeirantes.AnexoDiretorio** (nao referenciado no projeto Romaneio).

## Integracao em outras telas
1. Garantir que a tela saiba o `lote` e o `AUTONUM` de `TB_AG_CS` (ou LTL + lote).
2. Incluir na view um `Url.Action("Index", "ConsultaDocumentosAgendamento")` e montar a query string como acima, ou abrir com `window.open`.

## Pendencias / limitacoes
- **AnexoDiretorio**: exclusao/upload e obtencao via API externa nao foram portados; apenas leitura direta do blob na tabela.
- **Oracle**: o WinForms original mistura trechos Oracle (`TO_CHAR`, `SYSDATE`); o Romaneio usa **SQL Server** (`CONVERT`, `GETDATE()`). Ambientes Oracle exigiriam outra implementacao ou linked server.
- **Uniao CNTR**: parte das unioes do legado referencia `TB_CNTR_BL` com o mesmo parametro de agendamento usado no fluxo conteiner; no fluxo puramente CS da web essas linhas tendem a retornar vazio, como no legado quando o parametro e `TB_AG_CS.AUTONUM`.
