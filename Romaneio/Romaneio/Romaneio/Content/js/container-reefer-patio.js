(function () {
    var state = {
        detalhe: null,
        salvando: false
    };

    function alerta(msg, tipo) {
        tipo = tipo || 'info';
        var cls = tipo === 'danger' ? 'alert-danger' : tipo === 'success' ? 'alert-success' : 'alert-info';
        var html = '<div class="alert ' + cls + ' alert-dismissible fade in"><button type="button" class="close" data-dismiss="alert">&times;</button>' + (msg || '') + '</div>';
        $('#alert-container').html(html);
    }

    function esc(s) {
        if (s == null || s === undefined) return '';
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function limpar() {
        state.detalhe = null;
        $('#txtFinal,#txtIdCompleto').val('');
        $('#txtMin,#txtMax,#txtSetPoint,#txtUmidadeSet,#txtVentilacaoSet,#txtGateIn').val('');
        $('#txtTempSup,#txtTempRet,#txtEscala,#txtUmidade,#txtVentilacao').val('');
        $('#badges-reefer').empty();
        $('#tblHistorico tbody').empty();
    }

    function badges(d) {
        var b = [];
        if (d.ALERTA_REEFER_DESLIGADO) b.push('<span class="label label-danger">REEFER DESLIGADO</span>');
        if (d.PLUG_OFF_TEXTO) b.push('<span class="label label-warning">' + esc(d.PLUG_OFF_TEXTO) + '</span>');
        $('#badges-reefer').html(b.join(' '));
    }

    function preencherDetalhe(d) {
        state.detalhe = d;
        $('#txtIdCompleto').val(d.ID_CONTEINER || '');
        $('#txtMin').val(d.TXT_MIN || '');
        $('#txtMax').val(d.TXT_MAX || '');
        $('#txtSetPoint').val(d.TXT_SET_POINT || '');
        $('#txtUmidadeSet').val(d.TXT_UMIDADE_SET || '');
        $('#txtVentilacaoSet').val(d.TXT_VENTILACAO_SET || '');
        $('#txtGateIn').val(d.DATA_ENT_TEMP || '');
        $('#txtEscala').val(d.TXT_ESCALA || '');
        $('#txtTempSup,#txtTempRet,#txtUmidade,#txtVentilacao').val('');
        badges(d);
    }

    function carregarHistorico() {
        var d = state.detalhe;
        if (!d) return;
        $.post(crpUrls.historico, {
            autonumIpa: d.AUTONUM_IPA || 0,
            autonumRdx: d.AUTONUM_RDX || 0,
            autonumOp: d.AUTONUM_OP || 0
        }).done(function (r) {
            var tb = $('#tblHistorico tbody');
            tb.empty();
            if (!r.success) {
                alerta(r.message || 'Erro historico', 'danger');
                return;
            }
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + esc(x.DT_MONITORING) + '</td><td>' + esc(x.TEMP_SUP) + '</td><td>' + esc(x.TEMP_RET) + '</td><td>' + esc(x.UMIDADE) + '</td><td>' + esc(x.VENTILACAO) + '</td><td>' + esc(x.USUARIO) + '</td></tr>');
            });
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function carregarConteiner() {
        var id = ($('#txtIdCompleto').val() || '').trim();
        if (!id) {
            alerta('Informe o ID do conteiner', 'danger');
            return;
        }
        $.post(crpUrls.carregar, { idConteiner: id }).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            preencherDetalhe(r.data);
            carregarHistorico();
            alerta('Conteiner carregado', 'success');
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function buscarFinal() {
        var f = ($('#txtFinal').val() || '').trim();
        if (!f) {
            alerta('Informe o final', 'danger');
            return;
        }
        $.post(crpUrls.buscarFinal, { final: f }).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            $('#txtIdCompleto').val(r.idConteiner || '');
            carregarConteiner();
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function montarPayloadSalvar(confirmarDiv) {
        var d = state.detalhe;
        if (!d) return null;
        return {
            ID_CONTEINER: d.ID_CONTEINER,
            AUTONUM_IPA: d.AUTONUM_IPA || 0,
            AUTONUM_RDX: d.AUTONUM_RDX || 0,
            AUTONUM_OP: d.AUTONUM_OP || 0,
            TXT_SET_POINT: $('#txtSetPoint').val(),
            TXT_UMIDADE_SET: $('#txtUmidadeSet').val(),
            TXT_VENTILACAO_SET: $('#txtVentilacaoSet').val(),
            TXT_TEMP_SUP: $('#txtTempSup').val(),
            TXT_TEMP_RET: $('#txtTempRet').val(),
            TXT_UMIDADE: $('#txtUmidade').val(),
            TXT_VENTILACAO: $('#txtVentilacao').val(),
            TXT_ESCALA: $('#txtEscala').val(),
            CONFIRMAR_DIVERGENCIA_TEMPERATURA: !!confirmarDiv
        };
    }

    function salvar(confirmarDiv) {
        if (state.salvando) return;
        var p = montarPayloadSalvar(confirmarDiv);
        if (!p) {
            alerta('Carregue o conteiner antes', 'danger');
            return;
        }
        state.salvando = true;
        $.post(crpUrls.salvar, p).done(function (r) {
            state.salvando = false;
            if (r.exigeConfirmacaoDivergenciaTemperatura) {
                $('#modalConfirmTempBody').text(r.message || '');
                $('#modalConfirmTemp').modal('show');
                return;
            }
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            alerta(r.message || 'OK', 'success');
            carregarConteiner();
        }).fail(function () {
            state.salvando = false;
            alerta('Falha de rede', 'danger');
        });
    }

    function registrarPlugOff() {
        var d = state.detalhe;
        if (!d || !(d.AUTONUM_IPA > 0)) {
            alerta('Plug OFF disponivel apenas para unidade IPA carregada', 'danger');
            return;
        }
        if (d.ALERTA_REEFER_DESLIGADO) {
            alerta('Operacao nao permitida enquanto reefer desligado estiver ativo', 'danger');
            return;
        }
        $('#modalPlugOff').modal('show');
    }

    function executarPlugOff() {
        var d = state.detalhe;
        $.post(crpUrls.plugOff, { autonumIpa: d.AUTONUM_IPA }).done(function (r) {
            $('#modalPlugOff').modal('hide');
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            alerta(r.message || 'OK', 'success');
            carregarConteiner();
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function carregarEntradas() {
        $.post(crpUrls.entradas).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro entradas', 'danger');
                return;
            }
            $('#lblTotEnt').text(r.total != null ? r.total : (r.dados || []).length);
            var tb = $('#tblEntradas tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + esc(x.ID_CONTEINER) + '</td><td>' + esc(x.TAMANHO) + '</td><td>' + esc(x.SETPOINT) + '</td><td>' + esc(x.DT_ATRACACAO) + '</td><td>' + esc(x.ARM_ATRACACAO) + '</td><td>' + esc(x.ENTRADA_PREVISTA) + '</td></tr>');
            });
        });
    }

    function carregarEstoque() {
        $.post(crpUrls.estoque, {
            apenasDesligados: $('#chkApenasDesligados').is(':checked'),
            comAgendamentoSaida: $('#chkAgendSaida').is(':checked'),
            posicionados: $('#chkPosicionados').is(':checked')
        }).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro estoque', 'danger');
                return;
            }
            $('#lblTotSai').text(r.total != null ? r.total : (r.dados || []).length);
            var tb = $('#tblSaidas tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                var cls = (x.TEM_PLUG_OFF && Number(x.TEM_PLUG_OFF) > 0) ? ' class="warning"' : '';
                tb.append('<tr' + cls + '><td>' + esc(x.ID_CONTEINER) + '</td><td>' + esc(x.DT_ENTRADA) + '</td><td>' + esc(x.YARD) + '</td><td>' + esc(x.PERIODO_INICIAL) + '</td><td>' + esc(x.SETPOINT) + '</td><td>' + esc(x.TEMP_RET_M1) + '</td><td>' + esc(x.DT1) + '</td><td>' + esc(x.TEMP_RET_M2) + '</td><td>' + esc(x.DT2) + '</td></tr>');
            });
        });
    }

    function carregarDesligados() {
        $.post(crpUrls.desligados).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            $('#lblTotDesl').text(r.total != null ? r.total : (r.dados || []).length);
            var tb = $('#tblDesligados tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + esc(x.ID_CONTEINER) + '</td><td>' + esc(x.DT_ENTRADA) + '</td><td>' + esc(x.YARD) + '</td><td>' + esc(x.PERIODO_INICIAL) + '</td><td>' + esc(x.SETPOINT) + '</td><td>' + esc(x.TEMP_RET_M1) + '</td><td>' + esc(x.DT1) + '</td><td>' + esc(x.TEMP_RET_M2) + '</td><td>' + esc(x.DT2) + '</td><td>' + esc(x.DT_PLUG_OFF_FMT) + '</td></tr>');
            });
        });
    }

    $(function () {
        $('#btnBuscarFinal').click(buscarFinal);
        $('#btnCarregar').click(carregarConteiner);
        $('#btnLimpar').click(limpar);
        $('#btnSalvar').click(function () { salvar(false); });
        $('#btnConfirmarTemp').click(function () {
            $('#modalConfirmTemp').modal('hide');
            salvar(true);
        });
        $('#btnPlugOff').click(registrarPlugOff);
        $('#btnConfirmarPlugOff').click(executarPlugOff);
        $('#btnFiltrarEstoque').click(carregarEstoque);
        $('#btnAtualizarDesligados').click(carregarDesligados);

        $('#txtFinal').on('keypress', function (e) {
            if (e.which === 13) { e.preventDefault(); buscarFinal(); }
        });
        $('#txtIdCompleto').on('keypress', function (e) {
            if (e.which === 13) { e.preventDefault(); carregarConteiner(); }
        });

        $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
            var href = $(e.target).attr('href');
            if (href === '#tab-prev') {
                carregarEntradas();
                carregarEstoque();
            } else if (href === '#tab-desl') {
                carregarDesligados();
            }
        });
    });
})();
