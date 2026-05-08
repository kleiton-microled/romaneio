(function () {
    var config = window.carregamentoCsConfig || {};
    var urls = config.urls || {};
    var bipagem = config.bipagem || { obrigatoria: false, bloquearAposValidar: false };

    var estado = {
        placa: "",
        ordemSelecionada: null,
        marcanteSelecionado: null,
        marcanteValidado: false,
        ultimaTeclaTimestamp: 0,
        marcanteBuffer: ""
    };

    // ---------- Helpers de UI ----------

    function alerta(mensagem, tipo) {
        var cls = "alert-" + (tipo || "danger");
        $("#alert-container").html('<div class="alert ' + cls + '">' + mensagem + '</div>');
    }

    function limparAlerta() {
        $("#alert-container").empty();
    }

    function critica(mensagem, tipo) {
        if (!mensagem) {
            $("#txtCritica").hide().empty();
            return;
        }

        $("#txtCritica")
            .removeClass("alert-info alert-success alert-warning alert-danger")
            .addClass("alert-" + (tipo || "info"))
            .text(mensagem)
            .show();
    }

    function loadingAbrir(titulo) {
        if (!window.Swal || typeof window.Swal.fire !== "function") {
            return;
        }
        window.Swal.fire({
            title: titulo,
            allowOutsideClick: false,
            allowEscapeKey: false,
            didOpen: function () { window.Swal.showLoading(); }
        });
    }

    function loadingFechar() {
        if (window.Swal && typeof window.Swal.close === "function") {
            window.Swal.close();
        }
    }

    function ajaxGet(url, data, onSuccess) {
        $.ajax({
            url: url,
            type: "GET",
            data: data,
            success: onSuccess,
            error: function () { alerta("Erro ao comunicar com o servidor", "danger"); }
        });
    }

    function ajaxPost(url, data, onSuccess) {
        $.ajax({
            url: url,
            type: "POST",
            data: data,
            success: onSuccess,
            error: function () { alerta("Erro ao comunicar com o servidor", "danger"); }
        });
    }

    // ---------- Render ----------

    function renderOrdens(ordens) {
        var tbody = $("#tblOrdens tbody");
        tbody.empty();

        if (!ordens || ordens.length === 0) {
            tbody.append('<tr><td colspan="6" class="text-center">Nenhuma ordem para o veiculo</td></tr>');
            return;
        }

        $.each(ordens, function (_, ordem) {
            var classe = (ordem.QUANTIDADE > 0 && ordem.QUANTIDADE === ordem.QTDE_CARREGADA) ? "success" : "";
            var tr = $("<tr></tr>")
                .attr("data-ordem", JSON.stringify(ordem))
                .addClass(classe);
            tr.append("<td>" + ordem.NUM_OC + "</td>");
            tr.append("<td>" + ordem.LOTE + "</td>");
            tr.append("<td>" + ordem.ITEM + "</td>");
            tr.append("<td>" + ordem.QUANTIDADE + "</td>");
            tr.append("<td>" + ordem.QTDE_CARREGADA + "</td>");
            tr.append("<td>" + (ordem.EMBALAGEM || "") + "</td>");
            tbody.append(tr);
        });
    }

    function renderCarregamentos(marcantes) {
        var tbody = $("#tblCarregamentos tbody");
        tbody.empty();

        if (!marcantes || marcantes.length === 0) {
            tbody.append('<tr><td colspan="3" class="text-center">Nenhum carregamento</td></tr>');
            return;
        }

        $.each(marcantes, function (_, marcante) {
            var tr = $("<tr></tr>").attr("data-marcante", JSON.stringify(marcante));
            tr.append("<td>" + marcante.MARCANTE_FORMATADO + "</td>");
            tr.append("<td>" + marcante.QTDE + "</td>");
            tr.append('<td><button type="button" class="btn btn-xs btn-link js-selecionar-marcante">Selecionar</button></td>');
            tbody.append(tr);
        });
    }

    function renderLocais(locais) {
        var tbody = $("#tblLocais tbody");
        tbody.empty();

        if (!locais || locais.length === 0) {
            tbody.append('<tr><td colspan="3" class="text-center">Nenhuma localizacao disponivel</td></tr>');
            return;
        }

        $.each(locais, function (_, local) {
            var tr = $("<tr></tr>");
            tr.append("<td>" + local.MARCANTE_FORMATADO + "</td>");
            tr.append("<td>" + local.QTDE + "</td>");
            tr.append("<td>" + (local.LOCAL_COMPLETO || "") + "</td>");
            tbody.append(tr);
        });
    }

    function preencherCamposMarcante(dados) {
        if (!dados) {
            return;
        }
        $("#txtLote").val(dados.LOTE || "");
        $("#txtQtde").val(dados.VOLUMES || "");
        $("#txtArmazem").val(dados.DESCR_ARMAZEM || "");
        $("#txtLocal").val(dados.POSICAO_YARD || "");
    }

    function limparCamposMarcante() {
        $("#txtMarcante").val("");
        $("#txtLote").val("");
        $("#txtQtde").val("");
        $("#txtArmazem").val("");
        $("#txtLocal").val("");
        critica("");
        estado.marcanteSelecionado = null;
        estado.marcanteValidado = false;
        estado.marcanteBuffer = "";
        $("#btnCarregar, #btnCancelar").prop("disabled", true);
    }

    function limparTela() {
        $("#tblOrdens tbody").empty();
        $("#tblCarregamentos tbody").empty();
        estado.placa = "";
        estado.ordemSelecionada = null;
        $("#ddlVeiculo").val("");
        limparCamposMarcante();
        $("#btnDocumentos").prop("disabled", true);
    }

    // ---------- Acoes ----------

    function carregarOrdens() {
        var placa = $("#ddlVeiculo").val();
        estado.placa = placa || "";
        estado.ordemSelecionada = null;
        $("#tblCarregamentos tbody").empty();
        $("#btnDocumentos").prop("disabled", true);
        limparCamposMarcante();
        limparAlerta();

        if (!placa) {
            $("#tblOrdens tbody").empty();
            return;
        }

        loadingAbrir("Consultando ordens...");
        ajaxGet(urls.listarOrdens, { placa: placa }, function (res) {
            loadingFechar();
            if (!res.success) {
                alerta(res.message || "Falha ao listar ordens", "danger");
                return;
            }
            renderOrdens(res.ordens);
        });
    }

    function selecionarOrdem(tr) {
        $("#tblOrdens tbody tr").removeClass("info");
        tr.addClass("info");

        try {
            estado.ordemSelecionada = JSON.parse(tr.attr("data-ordem"));
        } catch (e) {
            estado.ordemSelecionada = null;
        }

        $("#btnDocumentos").prop("disabled", !(estado.ordemSelecionada && estado.ordemSelecionada.LOTE > 0));

        limparCamposMarcante();

        if (!estado.ordemSelecionada || !estado.ordemSelecionada.AUTONUMCS) {
            $("#tblCarregamentos tbody").empty();
            return;
        }

        ajaxGet(
            urls.listarCarregamentos,
            { placa: estado.placa, autonumCs: estado.ordemSelecionada.AUTONUMCS },
            function (res) {
                if (!res.success) {
                    alerta(res.message || "Falha ao listar carregamentos", "danger");
                    return;
                }
                renderCarregamentos(res.marcantes);
            }
        );
    }

    function selecionarMarcanteCarregado(tr) {
        try {
            var marcante = JSON.parse(tr.attr("data-marcante"));
            $("#txtMarcante").val(marcante.MARCANTE_FORMATADO);
            buscarMarcante(marcante.MARCANTE);
        } catch (e) { }
    }

    function buscarMarcante(numero) {
        if (!numero) {
            return;
        }

        ajaxGet(urls.buscarMarcante, { marcante: numero }, function (res) {
            if (!res.success) {
                alerta(res.message || "Marcante nao localizado", "danger");
                limparCamposMarcante();
                $("#txtMarcante").focus();
                return;
            }

            estado.marcanteSelecionado = res.dados;
            preencherCamposMarcante(res.dados);

            var jaCarregado = (res.dados.POSICAO_YARD || "").toUpperCase() === "CAM";
            var temOC = !!(estado.ordemSelecionada && estado.ordemSelecionada.AUTONUMCS);
            var divergente = temOC && estado.ordemSelecionada.AUTONUMCS !== res.dados.AUTONUMCS;

            if (divergente) {
                critica("Carga Divergente da Ordem de carregamento", "danger");
                $("#btnCarregar, #btnCancelar").prop("disabled", true);
                estado.marcanteValidado = false;
                return;
            }

            if (jaCarregado) {
                critica("Carga ja carregada - apenas cancelamento permitido", "warning");
                $("#btnCarregar").prop("disabled", true);
                $("#btnCancelar").prop("disabled", false);
            } else {
                critica("");
                $("#btnCarregar").prop("disabled", !temOC);
                $("#btnCancelar").prop("disabled", true);
            }

            estado.marcanteValidado = true;

            if (bipagem.obrigatoria && bipagem.bloquearAposValidar) {
                $("#txtMarcante").prop("readonly", true);
            }
        });
    }

    function abrirLocais() {
        if (!estado.ordemSelecionada || !estado.ordemSelecionada.AUTONUMCS) {
            alerta("Selecione uma ordem de carregamento", "warning");
            return;
        }

        ajaxGet(urls.listarLocais, { autonumCs: estado.ordemSelecionada.AUTONUMCS }, function (res) {
            if (!res.success) {
                alerta(res.message || "Falha ao listar locais", "danger");
                return;
            }
            renderLocais(res.locais);
            $("#modalLocais").modal("show");
        });
    }

    function abrirDocumento() {
        if (!estado.ordemSelecionada || !estado.ordemSelecionada.LOTE) {
            alerta("Selecione uma ordem para consultar documentos", "warning");
            return;
        }

        ajaxGet(urls.documento, { loteBl: estado.ordemSelecionada.LOTE }, function (res) {
            if (!res.success) {
                alerta(res.message || "Documento indisponivel", "warning");
                return;
            }

            var detalhe = res.ltl
                ? "Documento LTL para o lote " + res.loteBl
                : "Agendamento " + res.autonumAgCs + " para o lote " + res.loteBl;

            alerta(
                "Funcionalidade de impressao depende de integracao externa (legado VB6 chamava executavel local). "
                + detalhe + " - validar na homologacao tecnica.",
                "info"
            );
        });
    }

    function carregarMarcante() {
        if (!estado.marcanteSelecionado || !estado.ordemSelecionada) {
            alerta("Marcante e ordem precisam estar selecionados", "warning");
            return;
        }

        var qtde = parseInt($("#txtQtde").val(), 10);
        if (!qtde || qtde <= 0) {
            alerta("Quantidade indisponivel", "warning");
            return;
        }

        if (qtde > estado.marcanteSelecionado.VOLUMES) {
            alerta("Quantidade indisponivel", "warning");
            return;
        }

        loadingAbrir("Carregando marcante...");

        ajaxPost(urls.carregar, {
            AUTONUMCS_OC: estado.ordemSelecionada.AUTONUMCS,
            ORDEM: estado.ordemSelecionada.ORDEM_CARREG,
            PLACA_C: estado.placa,
            MARCANTE: estado.marcanteSelecionado.MARCANTE,
            QTDE: qtde,
            AUTONUM_CS_YARD: estado.marcanteSelecionado.AUTONUM_CS_YARD,
            AUTONUM_ARMAZEM: estado.marcanteSelecionado.AUTONUM_ARMAZEM
        }, function (res) {
            loadingFechar();

            if (!res.success) {
                critica(res.message || "Erro ao carregar", "danger");
                return;
            }

            alerta(res.message || "Carregado com sucesso", "success");
            critica("");
            limparCamposMarcante();
            $("#txtMarcante").prop("readonly", false).focus();
            carregarOrdens();
        });
    }

    function cancelarMarcante() {
        if (!estado.marcanteSelecionado) {
            alerta("Marcante nao informado/invalido", "warning");
            return;
        }

        var qtde = parseInt($("#txtQtde").val(), 10);
        if (!qtde || qtde <= 0) {
            qtde = estado.marcanteSelecionado.VOLUMES;
        }

        var confirmar = function () { executarCancelamento(qtde); };

        if (window.Swal && typeof window.Swal.fire === "function") {
            window.Swal.fire({
                icon: "warning",
                title: "Cancelar carregamento?",
                text: "O marcante sera devolvido ao estoque do armazem.",
                showCancelButton: true,
                confirmButtonText: "Cancelar carregamento",
                cancelButtonText: "Voltar"
            }).then(function (r) {
                if (r && r.isConfirmed) {
                    confirmar();
                }
            });
        } else if (window.confirm("Cancelar o carregamento deste marcante?")) {
            confirmar();
        }
    }

    function executarCancelamento(qtde) {
        loadingAbrir("Cancelando carregamento...");

        ajaxPost(urls.cancelar, {
            MARCANTE: estado.marcanteSelecionado.MARCANTE,
            AUTONUMCS_OC: estado.marcanteSelecionado.AUTONUMCS,
            AUTONUM_CS_YARD: estado.marcanteSelecionado.AUTONUM_CS_YARD,
            AUTONUM_ARMAZEM: estado.marcanteSelecionado.AUTONUM_ARMAZEM,
            QTDE: qtde
        }, function (res) {
            loadingFechar();

            if (!res.success) {
                critica(res.message || "Erro ao cancelar", "danger");
                return;
            }

            alerta(res.message || "Cancelado com sucesso", "success");
            critica("");
            limparCamposMarcante();
            $("#txtMarcante").prop("readonly", false).focus();
            carregarOrdens();
        });
    }

    function atualizarVeiculos() {
        loadingAbrir("Atualizando veiculos...");

        ajaxGet(urls.listarVeiculos, {}, function (res) {
            loadingFechar();

            if (!res.success) {
                alerta(res.message || "Falha ao atualizar veiculos", "danger");
                return;
            }

            var ddl = $("#ddlVeiculo");
            var anterior = ddl.val();
            ddl.empty().append('<option value="">Selecione</option>');
            $.each(res.veiculos, function (_, v) {
                ddl.append('<option value="' + v.PLACA_C + '">' + v.DISPLAY + "</option>");
            });
            ddl.val(anterior || "");
        });
    }

    // ---------- Anti-digitacao manual ----------

    function configurarAntiDigitacaoMarcante() {
        var $input = $("#txtMarcante");

        $input.on("keydown", function (event) {
            estado.ultimaTeclaTimestamp = event.timeStamp;
        });

        $input.on("paste", function (event) {
            if (bipagem.obrigatoria) {
                event.preventDefault();
                alerta("Bipagem obrigatoria - colar nao e permitido neste campo", "warning");
            }
        });

        $input.on("input", function () {
            this.value = (this.value || "").replace(/\D/g, "").substring(0, 9);
            estado.marcanteBuffer = this.value;
        });

        $input.on("keypress", function (event) {
            if (event.which === 13) {
                event.preventDefault();
                var numero = parseInt($input.val(), 10);
                if (!numero || isNaN(numero)) {
                    alerta("Marcante nao informado/invalido", "warning");
                    return;
                }
                buscarMarcante(numero);
            }
        });
    }

    // ---------- Atalhos ----------

    function configurarAtalhos() {
        $(document).on("keydown", function (event) {
            switch (event.which) {
                case 112: // F1
                    event.preventDefault();
                    abrirLocais();
                    break;
                case 113: // F2
                    event.preventDefault();
                    if (!$("#btnDocumentos").prop("disabled")) {
                        abrirDocumento();
                    }
                    break;
                case 114: // F3
                    event.preventDefault();
                    limparCamposMarcante();
                    $("#txtMarcante").prop("readonly", false).focus();
                    break;
                case 115: // F4
                    event.preventDefault();
                    if (!$("#btnCarregar").prop("disabled")) {
                        carregarMarcante();
                    }
                    break;
                case 116: // F5
                    event.preventDefault();
                    if (!$("#btnCancelar").prop("disabled")) {
                        cancelarMarcante();
                    }
                    break;
                case 27: // ESC
                    event.preventDefault();
                    window.location.href = urls.sair;
                    break;
            }
        });
    }

    // ---------- Bind ----------

    function bind() {
        $("#ddlVeiculo").on("change", carregarOrdens);
        $("#btnAtualizarVeiculos").on("click", atualizarVeiculos);

        $("#tblOrdens tbody").on("click", "tr", function () {
            if ($(this).find("td").length === 1) { return; }
            selecionarOrdem($(this));
        });

        $("#tblCarregamentos tbody").on("click", ".js-selecionar-marcante", function () {
            selecionarMarcanteCarregado($(this).closest("tr"));
        });

        $("#btnLocais").on("click", abrirLocais);
        $("#btnDocumentos").on("click", abrirDocumento);
        $("#btnLimpar").on("click", function () {
            limparCamposMarcante();
            $("#txtMarcante").prop("readonly", false).focus();
        });
        $("#btnCarregar").on("click", carregarMarcante);
        $("#btnCancelar").on("click", cancelarMarcante);
        $("#btnSair").on("click", function () { window.location.href = urls.sair; });

        configurarAntiDigitacaoMarcante();
        configurarAtalhos();
    }

    $(function () {
        bind();
        limparTela();
        $("#ddlVeiculo").focus();
    });
})();
