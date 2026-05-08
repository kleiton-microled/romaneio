(function () {
    var config = window.consultaLibCarregamentoConfig || {};
    var urls = config.urls || {};
    var estado = {
        dados: null
    };

    function alerta(mensagem, sucesso) {
        var cls = sucesso ? "alert-success" : "alert-danger";
        $("#alert-container").html('<div class="alert ' + cls + '">' + mensagem + '</div>');
    }

    function limparAlerta() {
        $("#alert-container").empty();
    }

    function mostrarLoadingFiltro() {
        if (!window.Swal || typeof window.Swal.fire !== "function") {
            return;
        }

        window.Swal.fire({
            title: "Consultando protocolo...",
            text: "Aguarde o carregamento dos dados.",
            allowOutsideClick: false,
            allowEscapeKey: false,
            didOpen: function () {
                window.Swal.showLoading();
            }
        });
    }

    function esconderLoadingFiltro() {
        if (!window.Swal || typeof window.Swal.close !== "function") {
            return;
        }

        window.Swal.close();
    }

    function mostrarLoadingRegistro() {
        if (!window.Swal || typeof window.Swal.fire !== "function") {
            return;
        }

        window.Swal.fire({
            title: "Registrando liberacao...",
            text: "Aguarde a conclusao da operacao.",
            allowOutsideClick: false,
            allowEscapeKey: false,
            didOpen: function () {
                window.Swal.showLoading();
            }
        });
    }

    function esconderLoadingRegistro() {
        if (!window.Swal || typeof window.Swal.close !== "function") {
            return;
        }

        window.Swal.close();
    }

    function limparTela() {
        limparAlerta();
        estado.dados = null;
        $("#txtProtocolo,#txtPlaca,#txtCntr,#txtBlFiltro").val("");
        $("#txtPeriodo,#txtLote,#txtDocumento,#txtTransportadora,#txtMotorista,#txtPlacas,#txtLocalizacao,#txtSiscarga,#txtSefaz,#txtCpfMot,#txtProtocoloFormatado").val("");
        $("#ddlTipoVeiculo").val("0");
        $("#tblCargas tbody").empty();
        $("#lstSimultaneos").empty();
        $("#badgeBloqueioBl,#badgeBloqueioCntr,#badgeGlme").hide();
        $("#btnRegistrar,#btnDocumentos").prop("disabled", true);
        $("#txtPlaca").focus();
    }

    function render(dados) {
        estado.dados = dados;
        $("#txtPeriodo").val(dados.PERIODO || "");
        $("#txtLote").val(dados.LOTE || "");
        $("#txtDocumento").val(dados.NUM_DOCUMENTO || "");
        $("#txtTransportadora").val(dados.TRANSPORTADORA || "");
        $("#txtMotorista").val(dados.MOTORISTA || "");
        $("#txtPlacas").val(dados.PLACAS || "");
        $("#txtLocalizacao").val(dados.LOCALIZACAO || "");
        $("#txtSiscarga").val(dados.STATUS_SISCARGA || "");
        $("#txtSefaz").val(dados.STATUS_SEFAZ || "");
        $("#txtCpfMot").val(dados.CPF_MOTORISTA || "");
        $("#txtProtocoloFormatado").val(dados.PROTOCOLO || "");
        $("#ddlTipoVeiculo").val((dados.TIPO_CAMINHAO || 0).toString());

        var tbody = $("#tblCargas tbody");
        tbody.empty();
        $.each(dados.CARGAS || [], function (_, item) {
            tbody.append("<tr><td>" + (item.CARGA || "") + "</td></tr>");
        });

        var lista = $("#lstSimultaneos");
        lista.empty();
        $.each(dados.PROTOCOLOS_SIMULTANEOS || [], function (_, item) {
            lista.append('<li class="list-group-item">' + item + "</li>");
        });

        $("#badgeBloqueioBl").toggle(!!dados.BLOQUEIO_BL);
        $("#badgeBloqueioCntr").toggle(!!dados.BLOQUEIO_CNTR);
        $("#badgeGlme").toggle(!!dados.GLME_PENDENTE);

        $("#btnRegistrar").prop("disabled", !dados.PODE_REGISTRAR);
        $("#btnDocumentos").prop("disabled", !dados.SUCESSO);
    }

    function filtrar() {
        limparAlerta();
        mostrarLoadingFiltro();

        $.ajax({
            url: urls.filtrar,
            type: "POST",
            data: {
                PROTOCOLO: $("#txtProtocolo").val(),
                PLACA: $("#txtPlaca").val(),
                CNTR: $("#txtCntr").val(),
                BL: $("#txtBlFiltro").val()
            },
            success: function (res) {
                if (!res.success) {
                    alerta(res.message || "Falha ao consultar", false);
                    return;
                }
                render(res.dados);
                if (res.message) {
                    alerta(res.message, true);
                }
            },
            error: function () {
                alerta("Erro ao comunicar com o servidor", false);
            },
            complete: function () {
                esconderLoadingFiltro();
            }
        });
    }

    function registrar() {
        if (!estado.dados) {
            alerta("Carga nao selecionada", false);
            return;
        }

        limparAlerta();
        mostrarLoadingRegistro();
        var placas = ($("#txtPlacas").val() || "").split("-");
        var placaCavalo = (placas[0] || "").trim();
        var placaCarreta = (placas[1] || "").trim();

        $.ajax({
            url: urls.registrar,
            type: "POST",
            data: {
                ORDEM_CARREGAMENTO: estado.dados.ORDEM_CARREGAMENTO,
                AUTONUM_LOTE: estado.dados.AUTONUM_LOTE,
                AUTONUM_CNTR: estado.dados.AUTONUM_CNTR,
                AUTONUM_CS: estado.dados.AUTONUM_CS,
                TIPO_CAMINHAO: $("#ddlTipoVeiculo").val(),
                PLACA_CAVALO: placaCavalo,
                PLACA_CARRETA: placaCarreta,
                CPF_MOTORISTA: $("#txtCpfMot").val()
            },
            success: function (res) {
                if (!res.success) {
                    alerta(res.message || "Nao foi possivel registrar a saida da carga!", false);
                    return;
                }

                alerta(res.message || "Registrado com sucesso", true);
                limparTela();
            },
            error: function () {
                alerta("Erro ao registrar a saida da carga!", false);
            },
            complete: function () {
                esconderLoadingRegistro();
            }
        });
    }

    function abrirDocumentos() {
        alerta("Acao de documentos depende de integracao externa e deve ser validada na homologacao.", false);
    }

    function bind() {
        $("#btnFiltrar").on("click", filtrar);
        $("#btnLimpar").on("click", limparTela);
        $("#btnRegistrar").on("click", registrar);
        $("#btnDocumentos").on("click", abrirDocumentos);
        $("#btnSair").on("click", function () { window.location.href = urls.sair; });

        $("#txtProtocolo").on("input", function () {
            this.value = this.value.replace(/[^\d/]/g, "").substring(0, 11);
        });
        $("#txtPlaca").on("input", function () {
            this.value = this.value.toUpperCase().substring(0, 8);
        });
        $("#txtCntr").on("input", function () {
            this.value = this.value.toUpperCase().substring(0, 8);
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
    });
})();
