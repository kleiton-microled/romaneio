(function () {
    var config = window.localizacaoCarregamentosConfig || {};
    var urls = config.urls || {};

    var estado = {
        selecionado: null
    };

    function alerta(mensagem, sucesso) {
        var cls = sucesso ? "alert-success" : "alert-danger";
        $("#alert-container").html('<div class="alert ' + cls + '">' + mensagem + "</div>");
    }

    function limparAlerta() {
        $("#alert-container").empty();
    }

    function limparDetalhe() {
        estado.selecionado = null;
        $("#txtCavalo,#txtCarreta,#txtLibera,#txtGateIn,#txtCarga,#txtQuadra,#txtRua,#txtTransportadora,#txtResponsavel").val("");
        $("#btnSaida").prop("disabled", true);
        $("#tblLocalizacao tbody tr").removeClass("info");
    }

    function limparTela() {
        limparAlerta();
        limparDetalhe();
        $("#txtPlaca").val("");
        $("#tblLocalizacao tbody").empty();
        $("#lblTotEnt,#lblTotCntr,#lblTotCs").text("0");
        $("#txtPlaca").focus();
    }

    function renderResultado(dados) {
        var itens = (dados && dados.ITENS) || [];
        var tbody = $("#tblLocalizacao tbody");
        tbody.empty();

        if (itens.length === 0) {
            tbody.append('<tr><td colspan="10" class="text-center">Nenhum carregamento localizado</td></tr>');
        } else {
            $.each(itens, function (_, item) {
                var tr = $("<tr></tr>").attr("data-item", JSON.stringify(item));
                tr.append("<td>" + (item.PLACA || "") + "</td>");
                tr.append("<td>" + (item.CARRETA || "") + "</td>");
                tr.append("<td>" + (item.DT_LIB_ENT_CAM || "") + "</td>");
                tr.append("<td>" + (item.DT_GATE_IN || "") + "</td>");
                tr.append("<td>" + (item.CARGA || "") + "</td>");
                tr.append("<td>" + (item.QUADRA || "") + "</td>");
                tr.append("<td>" + (item.RUA || "") + "</td>");
                tr.append("<td>" + (item.NOMETRANSPORTADORA || "") + "</td>");
                tr.append("<td>" + (item.USUARIO || "") + "</td>");
                tr.append("<td>" + (item.FLAG_CNTR === 1 ? "CNTR" : "CS") + "</td>");
                tbody.append(tr);
            });
        }

        $("#lblTotEnt").text((dados && dados.TOTAL_ENTRADAS) || 0);
        $("#lblTotCntr").text((dados && dados.TOTAL_CNTR) || 0);
        $("#lblTotCs").text((dados && dados.TOTAL_CS) || 0);
        limparDetalhe();
    }

    function selecionarLinha(tr) {
        if (!tr || tr.find("td").length <= 1) {
            return;
        }

        $("#tblLocalizacao tbody tr").removeClass("info");
        tr.addClass("info");

        try {
            estado.selecionado = JSON.parse(tr.attr("data-item"));
        } catch (e) {
            estado.selecionado = null;
        }

        if (!estado.selecionado) {
            return;
        }

        $("#txtCavalo").val(estado.selecionado.PLACA || "");
        $("#txtCarreta").val(estado.selecionado.CARRETA || "");
        $("#txtLibera").val(estado.selecionado.DT_LIB_ENT_CAM || "");
        $("#txtGateIn").val(estado.selecionado.DT_GATE_IN || "");
        $("#txtCarga").val(estado.selecionado.CARGA || "");
        $("#txtQuadra").val(estado.selecionado.QUADRA || "");
        $("#txtRua").val(estado.selecionado.RUA || "");
        $("#txtTransportadora").val(estado.selecionado.NOMETRANSPORTADORA || "");
        $("#txtResponsavel").val(estado.selecionado.USUARIO || "");
        $("#btnSaida").prop("disabled", estado.selecionado.FLAG_CNTR !== 1 || !estado.selecionado.AUTONUM_CNTR);
    }

    function filtrar() {
        limparAlerta();
        $.ajax({
            url: urls.filtrar,
            type: "POST",
            data: { PLACA: $("#txtPlaca").val() },
            success: function (res) {
                if (!res.success) {
                    alerta(res.message || "Falha ao filtrar", false);
                    renderResultado({ ITENS: [], TOTAL_ENTRADAS: 0, TOTAL_CNTR: 0, TOTAL_CS: 0 });
                    return;
                }

                renderResultado(res.dados);
                if (res.message) {
                    alerta(res.message, true);
                }
            },
            error: function () {
                alerta("Erro ao comunicar com o servidor", false);
            }
        });
    }

    function moverSaida() {
        if (!estado.selecionado || estado.selecionado.FLAG_CNTR !== 1 || !estado.selecionado.AUTONUM_CNTR) {
            alerta("Selecione um conteiner valido", false);
            return;
        }

        if (!window.Swal || typeof window.Swal.fire !== "function") {
            enviarMoverSaida();
            return;
        }

        window.Swal.fire({
            icon: "question",
            title: "Confirma o carregamento da unidade?",
            showCancelButton: true,
            confirmButtonText: "Confirmar",
            cancelButtonText: "Cancelar"
        }).then(function (r) {
            if (r && r.isConfirmed) {
                enviarMoverSaida();
            }
        });
    }

    function enviarMoverSaida() {
        $.ajax({
            url: urls.moverSaida,
            type: "POST",
            data: {
                AUTONUM_CNTR: estado.selecionado.AUTONUM_CNTR,
                QUADRA_ATUAL: estado.selecionado.QUADRA,
                FLAG_CNTR: estado.selecionado.FLAG_CNTR
            },
            success: function (res) {
                if (!res.success) {
                    alerta(res.message || "Nao foi possivel movimentar para SAIDA", false);
                    return;
                }

                alerta(res.message || "Movimentado para SAIDA", true);
                filtrar();
            },
            error: function () {
                alerta("Erro ao comunicar com o servidor", false);
            }
        });
    }

    function bind() {
        $("#btnFiltrar").on("click", filtrar);
        $("#btnLimpar").on("click", limparTela);
        $("#btnSaida").on("click", moverSaida);
        $("#btnSair").on("click", function () { window.location.href = urls.sair; });

        $("#txtPlaca").on("input", function () {
            this.value = (this.value || "").toUpperCase().substring(0, 8);
        });

        $("#txtPlaca").on("keypress", function (event) {
            if (event.which === 13) {
                event.preventDefault();
                filtrar();
            }
        });

        $("#tblLocalizacao tbody").on("click", "tr", function () {
            selecionarLinha($(this));
        });

        $(document).on("keydown", function (event) {
            if (event.key === "Escape") {
                event.preventDefault();
                window.location.href = urls.sair;
            }
        });
    }

    $(function () {
        bind();
        limparTela();
        filtrar();
    });
})();
