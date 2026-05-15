(function () {
    'use strict';

    var cfg = window.inventarioCsArmazemConfig || {};
    var estado = {
        modoTag: 'C',
        flagExportacao: false,
        loteTag: '',
        itemChave: '',
        qtdeDisponivel: 0,
        qtdeMarcante: 0,
        yardTag: '',
        flagModoConteiner: false
    };

    function mostrarCritica(msg, sucesso) {
        var $el = $('#txtCritica');
        $el.text(msg || '');
        $el.show();
        $el.removeClass('alert-info alert-danger alert-success');
        if (sucesso) {
            $el.addClass('alert-success');
        } else if (msg) {
            $el.addClass('alert-danger');
        } else {
            $el.addClass('alert-info');
        }
    }

    function limparFormulario() {
        $('#txtMarcante').val('');
        $('#txtLote').val('');
        $('#txtMercadoria,#txtBL,#txtCliente,#txtEmbalagem,#txtMarca,#txtEntrada,#txtConteiner,#txtIMO,#txtDocCanal,#txtMOV,#txtNVOCC,#txtLocalAtual').val('');
        $('#txtQtde,#txtQtdeM,#txtQtdePos,#txtAutonumCsYard').val('');
        $('#ddlItem').empty();
        $('#txtLocalPos').val('');
        $('#ddlArmazem').val('');
        $('#ddlMotivo').val('');
        $('#txtPrateleira').val('');
        $('#lblExp').hide();
        $('#lblTituloStatus').text('INVENTARIO DE CARGA SOLTA').removeClass('text-danger');
        $('#painelConteiner').hide();
        estado = { modoTag: 'C', flagExportacao: false, loteTag: '', itemChave: '', qtdeDisponivel: 0, qtdeMarcante: 0, yardTag: '', flagModoConteiner: false };
        mostrarCritica('', false);
    }

    function carregarCombos() {
        $.getJSON(cfg.urls.combos)
            .done(function (r) {
                if (!r.success) return;
                var $a = $('#ddlArmazem').empty();
                $a.append($('<option>').val('').text('Selecione'));
                (r.armazens || []).forEach(function (x) {
                    $a.append($('<option>').val(x.AUTONUM).text(x.DISPLAY));
                });
                var $m = $('#ddlMotivo').empty();
                $m.append($('<option>').val('').text('Selecione'));
                (r.motivos || []).forEach(function (x) {
                    $m.append($('<option>').val(x.AUTONUM).text(x.DISPLAY));
                });
            });
    }

    function aplicarDados(d) {
        if (!d) return;
        estado.modoTag = d.MODO_TAG || 'C';
        estado.flagExportacao = !!d.FLAG_EXPORTACAO;
        estado.loteTag = d.LOTE_TAG || '';
        $('#lblTituloStatus').text(d.TITULO_STATUS || 'INVENTARIO DE CARGA SOLTA');
        if (d.PATIO_DIVERGENTE) {
            $('#lblTituloStatus').addClass('text-danger');
        }
        $('#lblExp').toggle(!!d.FLAG_EXPORTACAO);
        $('#txtLote').val(d.LOTE || '');
        $('#txtMercadoria').val(d.MERCADORIA || '');
        $('#txtBL').val(d.BL || '');
        $('#txtCliente').val(d.CLIENTE || '');
        $('#txtEmbalagem').val(d.EMBALAGEM || '');
        $('#txtMarca').val(d.MARCA || '');
        $('#txtEntrada').val(d.ENTRADA || '');
        $('#txtConteiner').val(d.CONTEINER || '');
        $('#txtIMO').val(d.IMO || '');
        $('#txtDocCanal').val(((d.DOC || '') + ' / ' + (d.CANAL || '')).trim());
        $('#txtMOV').val(d.MOV || '');
        $('#txtNVOCC').val(d.NVOCC || '');
        $('#txtLocalAtual').val(d.LOCAL_ATUAL || '');
        $('#txtQtde').val(d.QTDE || 0);
        $('#txtQtdeM').val(d.QTDE_M || 0);
        $('#txtQtdePos').val(d.QTDE_POS || 0);
        estado.qtdeDisponivel = d.QTDE || 0;
        estado.qtdeMarcante = d.QTDE_M || 0;
        $('#txtAutonumCsYard').val(d.AUTONUM_CS_YARD || 0);
        if (d.MOTIVO_PADRAO) {
            $('#ddlMotivo').val(String(d.MOTIVO_PADRAO));
        }
        var $it = $('#ddlItem').empty();
        $it.append($('<option>').val('').text('Selecione'));
        (d.ITENS || []).forEach(function (x) {
            $it.append($('<option>').val(x.ITEM_CHAVE).text(x.DISPLAY || x.ITEM_CHAVE));
        });
    }

    function carregarMarcante() {
        var m = parseInt($('#txtMarcante').val(), 10);
        if (!m) {
            limparFormulario();
            return;
        }
        $.getJSON(cfg.urls.marcante, { id: m })
            .done(function (r) {
                if (!r.success || !r.dados) {
                    mostrarCritica((r.dados && r.dados.MENSAGEM) || 'Marcante nao localizado', false);
                    return;
                }
                aplicarDados(r.dados);
                mostrarCritica('', false);
            })
            .fail(function () {
                mostrarCritica('Erro ao consultar marcante', false);
            });
    }

    function carregarItemDetalhe() {
        var id = $('#ddlItem').val();
        if (!id) return;
        $.getJSON(cfg.urls.item, { idGravacao: id })
            .done(function (r) {
                if (!r.success) return;
                $('#txtQtde').val(r.qtde);
                $('#txtEmbalagem').val(r.embalagem);
                estado.qtdeDisponivel = r.qtde;
                $('#txtQtdePos').val(r.qtde);
                if (r.motivoPadrao) {
                    $('#ddlMotivo').val(String(r.motivoPadrao));
                }
            });
    }

    function resolverPrateleiraServidor() {
        var t = ($('#txtPrateleira').val() || '').trim();
        if (!t) return;
        $.getJSON(cfg.urls.resolverPrateleira, { texto: t })
            .done(function (r) {
                if (!r.success) {
                    mostrarCritica(r.message || 'Etiqueta invalida', false);
                    return;
                }
                estado.flagModoConteiner = !!r.modoContener;
                if (r.modoContener) {
                    $('#painelConteiner').show();
                    $('#ddlArmazem').val(String(r.armazemId));
                    $('#txtLocalPos').val('').prop('readonly', true);
                } else {
                    $('#painelConteiner').hide();
                    $('#ddlArmazem').val(String(r.armazemId));
                    $('#txtLocalPos').val(r.yard || '').prop('readonly', false);
                    estado.yardTag = r.yardTag || '';
                }
            });
    }

    function salvar() {
        var req = {
            MARCANTE: parseInt($('#txtMarcante').val(), 10) || 0,
            LOTE: $('#txtLote').val() || '',
            LOTE_TAG: estado.loteTag,
            ITEM_CHAVE: $('#ddlItem').val() || '',
            MODO_TAG: estado.modoTag,
            FLAG_EXPORTACAO: estado.flagExportacao,
            FLAG_MODO_CONTENER: estado.flagModoConteiner,
            ARMAZEM_ID: parseInt($('#ddlArmazem').val(), 10) || 0,
            YARD: ($('#txtLocalPos').val() || '').trim(),
            YARD_TAG: estado.yardTag || '',
            QUANTIDADE_POS: parseInt($('#txtQtdePos').val(), 10) || 0,
            QTDE_DISPONIVEL: parseInt($('#txtQtde').val(), 10) || estado.qtdeDisponivel,
            QTDE_MARCANTE: parseInt($('#txtQtdeM').val(), 10) || estado.qtdeMarcante,
            MOTIVO_ID: parseInt($('#ddlMotivo').val(), 10) || 0,
            YARD_ATUAL_TEXTO: $('#txtLocalAtual').val() || '',
            FL_FRENTE: $('#chkFrente').is(':checked') ? 1 : 0,
            FL_FUNDO: $('#chkFundo').is(':checked') ? 1 : 0,
            FL_LE: $('#chkLe').is(':checked') ? 1 : 0,
            FL_LD: $('#chkLd').is(':checked') ? 1 : 0,
            PERCENT_CT: parseInt($('#txtPercentCt').val(), 10) || 0,
            IMO: $('#txtIMO').val() || ''
        };

        $.ajax({
            url: cfg.urls.salvar,
            type: 'POST',
            contentType: 'application/json; charset=utf-8',
            data: JSON.stringify(req)
        }).done(function (r) {
            if (r.success) {
                mostrarCritica(r.message || 'OK', true);
                limparFormulario();
                $('#txtMarcante').focus();
            } else {
                mostrarCritica(r.message || 'Erro', false);
            }
        }).fail(function (xhr) {
            var msg = 'Erro ao salvar';
            if (xhr.responseJSON && xhr.responseJSON.message) {
                msg = xhr.responseJSON.message;
            }
            mostrarCritica(msg, false);
        });
    }

    function abrirHistorico() {
        var m = parseInt($('#txtMarcante').val(), 10);
        if (!m) {
            Swal.fire('Informe o marcante', '', 'info');
            return;
        }
        $.getJSON(cfg.urls.historico, { marcante: m })
            .done(function (r) {
                var $tb = $('#tblHistorico tbody').empty();
                (r.linhas || []).forEach(function (l) {
                    $tb.append('<tr><td>' + (l.ARMAZEM || '') + '</td><td>' + (l.LOCAL || '') + '</td><td>' + (l.DATA_MOV || '') + '</td><td>' + (l.USUARIO || '') + '</td></tr>');
                });
                $('#modalHistorico').modal('show');
            });
    }

    $(function () {
        carregarCombos();

        $('#txtMarcante').on('blur', carregarMarcante);
        $('#ddlItem').on('change', carregarItemDetalhe);
        $('#txtPrateleira').on('blur', resolverPrateleiraServidor);

        $('#btnHistorico').on('click', abrirHistorico);
        $('#btnLimpar').on('click', function () {
            Swal.fire({
                title: 'Limpar dados?',
                showCancelButton: true,
                confirmButtonText: 'Sim',
                cancelButtonText: 'Nao'
            }).then(function (x) {
                if (x.isConfirmed) {
                    limparFormulario();
                }
            });
        });
        $('#btnSalvar').on('click', salvar);
        $('#btnSair').on('click', function () {
            window.location.href = cfg.urls.home;
        });

        $(document).on('keydown', function (e) {
            if (e.key === 'Escape') {
                if ($('#modalHistorico').hasClass('in')) {
                    $('#modalHistorico').modal('hide');
                } else {
                    window.location.href = cfg.urls.home;
                }
                return;
            }
            if (e.key === 'F1') {
                e.preventDefault();
                abrirHistorico();
            }
            if (e.key === 'F2') {
                e.preventDefault();
                limparFormulario();
            }
            if (e.key === 'F5') {
                e.preventDefault();
                salvar();
            }
        });
    });
})();
