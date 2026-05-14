(function () {
    var config = window.manutencaoMarcantesArmazemConfig || {};
    var urls = config.urls || {};

    var state = {
        autonumCs: null,
        quantidadeReferencia: null,
        embalagem: "",
        marcanteSelecionado: null
    };

    function showAlert(message, type) {
        var cssClass = type === "success" ? "alert-success" : "alert-danger";
        var html = '<div class="alert ' + cssClass + '" role="alert">' + message + "</div>";
        $("#alert-container").html(html);
    }

    function clearAlert() {
        $("#alert-container").empty();
    }

    function swalOk() {
        return window.Swal && typeof window.Swal.fire === "function";
    }

    function swalFiltroErro(message) {
        clearAlert();
        if (swalOk()) {
            window.Swal.fire({
                icon: "error",
                title: "Filtro do lote",
                text: message || "",
                confirmButtonText: "OK"
            });
            return;
        }
        showAlert(message, "error");
    }

    function swalFiltroAviso(message) {
        clearAlert();
        if (swalOk()) {
            window.Swal.fire({
                icon: "warning",
                title: "Filtro do lote",
                text: message || "",
                confirmButtonText: "OK"
            });
            return;
        }
        showAlert(message, "error");
    }

    function swalFiltroSucesso() {
        clearAlert();
        if (swalOk()) {
            window.Swal.fire({
                icon: "success",
                title: "Filtro aplicado",
                text: "Dados do lote carregados.",
                timer: 2200,
                showConfirmButton: true
            });
            return;
        }
    }

    function swalSalvarErro(message) {
        clearAlert();
        if (swalOk()) {
            window.Swal.fire({
                icon: "error",
                title: "Salvar",
                text: message || "",
                confirmButtonText: "OK"
            });
            return;
        }
        showAlert(message, "error");
    }

    function swalSalvarAviso(message) {
        clearAlert();
        if (swalOk()) {
            window.Swal.fire({
                icon: "warning",
                title: "Salvar",
                text: message || "",
                confirmButtonText: "OK"
            });
            return;
        }
        showAlert(message, "error");
    }

    function swalSalvarSucesso(message) {
        clearAlert();
        if (swalOk()) {
            window.Swal.fire({
                icon: "success",
                title: "Salvar",
                text: message || "Volumes atualizados com sucesso.",
                confirmButtonText: "OK"
            });
            return;
        }
        showAlert(message || "Volumes atualizados com sucesso.", "success");
    }

    function post(url, data, onSuccess, onAjaxError) {
        $.ajax({
            url: url,
            type: "POST",
            data: data,
            success: onSuccess,
            error: function () {
                if (typeof onAjaxError === "function") {
                    onAjaxError();
                    return;
                }
                swalFiltroErro("Erro ao comunicar com o servidor.");
            }
        });
    }

    function atualizarCorTotal(confere) {
        var $t = $("#txtTotalLista");
        $t.removeClass("text-success text-danger text-muted");
        if (state.quantidadeReferencia == null || state.quantidadeReferencia === "") {
            $t.addClass("text-muted");
            return;
        }
        if (confere) {
            $t.addClass("text-success");
        } else {
            $t.addClass("text-danger");
        }
    }

    function aplicarRespostaLista(response) {
        state.autonumCs = response.autonumCs;
        state.quantidadeReferencia = response.quantidadeReferencia;
        state.embalagem = response.embalagem || "";

        $("#txtEmbalagem").val(state.embalagem);
        $("#txtQtdeRef").val(
            response.quantidadeReferencia != null ? response.quantidadeReferencia : ""
        );
        $("#txtTotalLista").val(response.totalVolumesListados != null ? response.totalVolumesListados : "");
        atualizarCorTotal(!!response.totalConfereReferencia);

        var tbody = $("#tblMarcantes tbody");
        tbody.empty();
        state.marcanteSelecionado = null;
        $("#txtMarcanteSel").val("");
        $("#txtNovaQtde").val("");

        var linhas = response.linhas || [];
        $.each(linhas, function (_, item) {
            var tr = $("<tr></tr>");
            tr.attr("data-marcante", item.Marcante);
            tr.append("<td>" + item.Indice + "</td>");
            tr.append("<td>" + (item.DescrArmazem || "") + "</td>");
            tr.append("<td>" + (item.Posicao || "") + "</td>");
            tr.append("<td>" + (item.Item != null ? item.Item : "") + "</td>");
            tr.append("<td>" + item.Qtde + "</td>");
            tr.append("<td>" + item.MarcanteFormatado + "</td>");
            tbody.append(tr);
        });
    }

    function filtrar() {
        clearAlert();
        var lote = ($("#txtLote").val() || "").trim();
        if (!lote) {
            swalFiltroAviso("Informe o lote");
            return;
        }

        post(urls.filtrar, { lote: lote }, function (response) {
            if (!response.success) {
                swalFiltroErro(response.message || "Nao foi possivel filtrar o lote.");
                state.autonumCs = null;
                state.quantidadeReferencia = null;
                $("#txtEmbalagem").val("");
                $("#txtQtdeRef").val("");
                $("#txtTotalLista").val("");
                $("#tblMarcantes tbody").empty();
                atualizarCorTotal(false);
                return;
            }

            aplicarRespostaLista(response);
            swalFiltroSucesso();
        });
    }

    function salvar() {
        clearAlert();
        if (!state.autonumCs) {
            swalSalvarAviso("Informe o lote e filtre antes de salvar.");
            return;
        }

        if (!state.marcanteSelecionado) {
            swalSalvarAviso("Selecione o marcante");
            return;
        }

        var novaQtdeVal = $("#txtNovaQtde").val();
        if (novaQtdeVal === "" || novaQtdeVal === null) {
            swalSalvarAviso("Informe a quantidade");
            return;
        }

        post(
            urls.salvar,
            {
                AutonumCs: state.autonumCs,
                MarcanteAutonum: state.marcanteSelecionado,
                NovaQuantidade: novaQtdeVal,
                QuantidadeReferencia: state.quantidadeReferencia,
                Embalagem: state.embalagem
            },
            function (response) {
                if (!response.success) {
                    swalSalvarErro(response.message || "Nao foi possivel salvar.");
                    return;
                }

                aplicarRespostaLista(response);
                swalSalvarSucesso(response.message);
            },
            function () {
                swalSalvarErro("Erro ao comunicar com o servidor.");
            }
        );
    }

    function sair() {
        window.location.href = urls.sair;
    }

    function bindEvents() {
        $("#btnFiltrar").on("click", filtrar);
        $("#btnSalvar").on("click", salvar);
        $("#btnSair").on("click", sair);

        $("#txtLote").on("input", function () {
            this.value = this.value.replace(/\D/g, "").substring(0, 7);
        });

        $("#tblMarcantes").on("click", "tbody tr", function () {
            $("#tblMarcantes tbody tr").removeClass("info");
            $(this).addClass("info");
            state.marcanteSelecionado = $(this).data("marcante");
            var fmt = String(state.marcanteSelecionado || "");
            while (fmt.length < 9) {
                fmt = "0" + fmt;
            }
            $("#txtMarcanteSel").val(fmt);
        });

        $(document).on("keydown", function (event) {
            var tag = (event.target.tagName || "").toLowerCase();
            var isInput = tag === "input" || tag === "textarea" || tag === "select";

            if (event.key === "F4") {
                event.preventDefault();
                filtrar();
            } else if (event.key === "F5") {
                event.preventDefault();
                salvar();
            } else if (event.key === "Escape") {
                event.preventDefault();
                sair();
            }
        });
    }

    $(function () {
        bindEvents();
        $("#txtTotalLista").addClass("text-muted");
    });
})();
