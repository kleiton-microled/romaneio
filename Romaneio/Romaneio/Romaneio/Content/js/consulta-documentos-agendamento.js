(function () {
    var cfg = window.consultaDocumentosAgendaConfig;
    if (!cfg || !cfg.urls) {
        return;
    }

    var urls = cfg.urls;
    var ltl = !!cfg.ltl;
    var lote = cfg.lote || 0;
    var ag = cfg.ag || 0;

    function alerta(msg, ok) {
        var cls = ok ? "alert-success" : "alert-danger";
        $("#alert-container").html('<div class="alert ' + cls + '">' + msg + "</div>");
    }

    function limparPreview() {
        $("#ifrDoc").hide().attr("src", "about:blank");
        $("#imgDoc").hide().attr("src", "");
        $("#lblPreview").show().text("Selecione um documento na lista.");
    }

    function carregarLista() {
        if (!lote) {
            return;
        }

        $.ajax({
            url: urls.listar,
            type: "GET",
            data: { ltl: ltl, lote: lote, ag: ag },
            success: function (res) {
                if (!res.success) {
                    alerta(res.message || "Falha ao listar documentos", false);
                    return;
                }

                var tbody = $("#tblDocs tbody");
                tbody.empty();
                if (!res.itens || res.itens.length === 0) {
                    tbody.append('<tr><td colspan="3" class="text-center">Nenhum documento encontrado</td></tr>');
                    limparPreview();
                    return;
                }

                $.each(res.itens, function (_, it) {
                    var tr = $("<tr></tr>")
                        .attr("data-id", it.AUTONUM_AV_IMAGEM)
                        .css("cursor", "pointer");
                    tr.append("<td>" + (it.DESCRICAO || "") + "</td>");
                    tr.append("<td>" + (it.NOME_IMG || "") + "</td>");
                    tr.append("<td>" + (it.DT_INCLUSAO || "") + "</td>");
                    tbody.append(tr);
                });
            },
            error: function () {
                alerta("Erro ao comunicar com o servidor", false);
            }
        });
    }

    function abrirDocumento(id) {
        var q = $.param({ id: id, ltl: ltl, lote: lote, ag: ag });
        var url = urls.arquivo + "?" + q;

        $("#lblPreview").hide();

        $.ajax({
            url: url,
            type: "GET",
            xhrFields: { responseType: "blob" },
            success: function (data, status, xhr) {
                var ct = (xhr.getResponseHeader("Content-Type") || "").toLowerCase();
                var blobUrl = window.URL.createObjectURL(data);

                if (ct.indexOf("application/pdf") >= 0) {
                    $("#imgDoc").hide();
                    $("#ifrDoc").show().attr("src", blobUrl);
                } else if (ct.indexOf("image/") >= 0) {
                    $("#ifrDoc").hide().attr("src", "about:blank");
                    $("#imgDoc").show().attr("src", blobUrl);
                } else if (ct.indexOf("text/plain") >= 0) {
                    $("#ifrDoc").hide();
                    $("#imgDoc").hide();
                    var reader = new FileReader();
                    reader.onload = function () {
                        $("#lblPreview").show().text(reader.result);
                    };
                    reader.readAsText(data);
                } else {
                    $("#ifrDoc").hide();
                    $("#imgDoc").hide();
                    $("#lblPreview").show().text("Tipo de arquivo nao suportado para visualizacao inline. Use Salvar como no navegador se disponivel.");
                }
            },
            error: function (xhr) {
                limparPreview();
                var msg = "Nao foi possivel carregar o documento.";
                if (xhr && xhr.responseText) {
                    msg = xhr.responseText;
                }
                alerta(msg, false);
            }
        });
    }

    function fecharBloqueio(callback) {
        if (ltl || !ag) {
            if (callback) {
                callback();
            }
            return;
        }

        $.ajax({
            url: urls.fechar,
            type: "POST",
            data: { ag: ag },
            success: function (res) {
                if (callback) {
                    callback();
                }
            },
            error: function () {
                if (callback) {
                    callback();
                }
            }
        });
    }

    function bind() {
        $("#tblDocs tbody").on("click", "tr[data-id]", function () {
            $("#tblDocs tbody tr").removeClass("info");
            $(this).addClass("info");
            abrirDocumento($(this).attr("data-id"));
        });

        $("#btnFechar").on("click", function () {
            fecharBloqueio(function () {
                try {
                    window.close();
                } catch (e) { }
                if (window.opener) {
                    window.close();
                } else {
                    window.location.href = urls.home;
                }
            });
        });

        $("#btnVoltar").on("click", function () {
            fecharBloqueio(function () {
                window.location.href = urls.home;
            });
        });
    }

    $(function () {
        if ($("#tblDocs").length) {
            bind();
            carregarLista();
        }
    });
})();
