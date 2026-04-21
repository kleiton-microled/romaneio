(function () {
    var config = window.associacaoMarcantesConfig || {};
    var urls = config.urls || {};

    var state = {
        modoExp: false,
        itemSelecionado: null,
        marcanteSelecionado: null
    };

    function showAlert(message, type) {
        var cssClass = type === "success" ? "alert-success" : "alert-danger";
        var html = '<div class="alert ' + cssClass + '" role="alert">' + message + '</div>';
        $("#alert-container").html(html);
    }

    function clearAlert() {
        $("#alert-container").empty();
    }

    function post(url, data, onSuccess) {
        $.ajax({
            url: url,
            type: "POST",
            data: data,
            success: onSuccess,
            error: function () {
                showAlert("Erro ao comunicar com o servidor.", "error");
            }
        });
    }

    function carregarConteiners() {
        var lote = $("#txtLote").val();
        post(urls.obterConteiners, { lote: lote }, function (response) {
            if (!response.success) {
                showAlert(response.message, "error");
                return;
            }

            var ddl = $("#ddlCntr");
            ddl.empty();
            ddl.append('<option value="">Selecione</option>');
            $.each(response.conteineres, function (_, item) {
                ddl.append('<option value="' + item + '">' + item + '</option>');
            });
        });
    }

    function renderCargas(itens) {
        var tbody = $("#tblCargas tbody");
        tbody.empty();

        $.each(itens, function (_, item) {
            var tr = $("<tr></tr>");
            tr.attr("data-item", JSON.stringify(item));
            tr.append("<td>" + item.INDICE + "</td>");
            tr.append("<td>" + item.LOTE + "</td>");
            tr.append("<td>" + item.ITEM + "</td>");
            tr.append("<td>" + item.QUANTIDADE + "</td>");
            tr.append("<td>" + (item.EMBALAGEM || "") + "</td>");
            tr.append("<td>" + (item.ID_CONTEINER || "") + "</td>");
            tr.append("<td>" + (item.MERCADORIA || "") + "</td>");
            tbody.append(tr);
        });

        $("#btnAdicionar, #btnRemover").prop("disabled", itens.length === 0);
    }

    function renderTemporarios(itens) {
        var tbody = $("#tblMarcantesTemp tbody");
        tbody.empty();
        state.marcanteSelecionado = null;

        $.each(itens, function (_, item) {
            var tr = $("<tr></tr>");
            tr.attr("data-marcante", item.MARCANTE);
            tr.append("<td>" + item.MARCANTE_FORMATADO + "</td>");
            tr.append("<td>" + item.QTDE + "</td>");
            tbody.append(tr);
        });
    }

    function filtrar() {
        clearAlert();
        state.itemSelecionado = null;
        renderTemporarios([]);

        post(urls.filtrar, { lote: $("#txtLote").val(), cntr: $("#ddlCntr").val() }, function (response) {
            if (!response.success) {
                showAlert(response.message, "error");
                return;
            }

            state.modoExp = response.modoExp;
            $("#lblExp").toggle(!!state.modoExp);
            renderCargas(response.itens || []);
        });
    }

    function carregarTemporarios() {
        if (!state.itemSelecionado) {
            showAlert("Carga nao selecionada", "error");
            return;
        }

        post(urls.carregarTemporarios, {
            modoExp: state.modoExp,
            autonumRef: state.itemSelecionado.AUTONUMCS
        }, function (response) {
            if (!response.success) {
                showAlert(response.message, "error");
                return;
            }

            renderTemporarios(response.itens || []);
        });
    }

    function adicionar() {
        clearAlert();
        if (!state.itemSelecionado) {
            showAlert("Carga nao selecionada", "error");
            return;
        }

        post(urls.adicionarTemp, {
            MODO_EXP: state.modoExp,
            AUTONUM_REF: state.itemSelecionado.AUTONUMCS,
            MARCANTE: $("#txtMarcante").val(),
            QTDE: $("#txtQtde").val()
        }, function (response) {
            if (!response.success) {
                showAlert(response.message, "error");
                return;
            }

            $("#txtMarcante").val("").focus();
            $("#txtQtde").val("");
            renderTemporarios(response.itens || []);
        });
    }

    function remover() {
        clearAlert();
        if (!state.itemSelecionado) {
            showAlert("Carga nao selecionada", "error");
            return;
        }

        if (!state.marcanteSelecionado) {
            showAlert("Selecione um marcante para remover", "error");
            return;
        }

        post(urls.removerTemp, {
            MODO_EXP: state.modoExp,
            AUTONUM_REF: state.itemSelecionado.AUTONUMCS,
            MARCANTE: state.marcanteSelecionado
        }, function (response) {
            if (!response.success) {
                showAlert(response.message, "error");
                return;
            }

            renderTemporarios(response.itens || []);
        });
    }

    function salvar() {
        clearAlert();
        if (!state.itemSelecionado) {
            showAlert("Carga nao selecionada", "error");
            return;
        }

        post(urls.salvar, {
            MODO_EXP: state.modoExp,
            ITEM_SELECIONADO: state.itemSelecionado
        }, function (response) {
            if (!response.success) {
                showAlert(response.message, "error");
                return;
            }

            showAlert(response.message, "success");
            state.itemSelecionado = null;
            state.marcanteSelecionado = null;
            $("#txtLote").focus();
            $("#tblCargas tbody").empty();
            $("#tblMarcantesTemp tbody").empty();
            $("#lblExp").hide();
        });
    }

    function sair() {
        window.location.href = urls.sair;
    }

    function bindEvents() {
        $("#txtLote").on("blur", carregarConteiners);
        $("#btnFiltrar").on("click", filtrar);
        $("#btnAdicionar").on("click", adicionar);
        $("#btnRemover").on("click", remover);
        $("#btnSalvar").on("click", salvar);
        $("#btnSair").on("click", sair);

        $("#txtMarcante").on("input", function () {
            this.value = this.value.replace(/\D/g, "").substring(0, 9);
        });

        $("#txtQtde").on("input", function () {
            if (parseInt(this.value || "0", 10) < 0) {
                this.value = "";
            }
        });

        $("#tblCargas").on("click", "tbody tr", function () {
            $("#tblCargas tbody tr").removeClass("info");
            $(this).addClass("info");
            state.itemSelecionado = $(this).data("item");
            carregarTemporarios();
        });

        $("#tblMarcantesTemp").on("click", "tbody tr", function () {
            $("#tblMarcantesTemp tbody tr").removeClass("info");
            $(this).addClass("info");
            state.marcanteSelecionado = $(this).data("marcante");
        });

        $(document).on("keydown", function (event) {
            var tag = (event.target.tagName || "").toLowerCase();
            var isInput = tag === "input" || tag === "textarea" || tag === "select";

            if (event.key === "F1") {
                event.preventDefault();
                adicionar();
            } else if (event.key === "F2") {
                event.preventDefault();
                remover();
            } else if (event.key === "F4") {
                event.preventDefault();
                filtrar();
            } else if (event.key === "F5") {
                event.preventDefault();
                salvar();
            } else if (event.key === "Escape" && !isInput) {
                event.preventDefault();
                sair();
            }
        });
    }

    $(function () {
        bindEvents();
    });
})();
