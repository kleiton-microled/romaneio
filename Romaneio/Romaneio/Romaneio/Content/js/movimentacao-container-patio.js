(function () {
    var state = {
        detalhe: null,
        pendenteCamera: false,
        usuarioConfirmouSaida: false,
        usuarioConfirmouVeiculoServ: false,
        salvando: false
    };

    function alerta(msg, tipo) {
        tipo = tipo || 'info';
        var cls = tipo === 'danger' ? 'alert-danger' : tipo === 'success' ? 'alert-success' : 'alert-info';
        var html = '<div class="alert ' + cls + ' alert-dismissible fade in"><button type="button" class="close" data-dismiss="alert">&times;</button>' + (msg || '') + '</div>';
        $('#alert-container').html(html);
    }

    function limparFormMov() {
        state.detalhe = null;
        state.pendenteCamera = false;
        state.usuarioConfirmouSaida = false;
        state.usuarioConfirmouVeiculoServ = false;
        $('#txtFinal').val('');
        $('#txtIdCompleto').val('');
        $('#txtYardAtual').val('');
        $('#txtDestino').val('');
        $('#ddlMotivo').val('');
        $('#chkLacresConf').prop('checked', false);
        $('#txtLacreIpa').val('');
        $('#txtLacre1').val('');
        $('#txtLacre2').val('');
        $('#txtLacre3').val('');
        $('#txtLacreLona').val('');
        $('#txtTara').val('');
        $('#txtContexto').val('');
        $('#badges').empty();
        $('#tblProximos tbody').empty();
        $('#tblVizinhos tbody').empty();
        limparTabDados();
    }

    function limparTabDados() {
        $('#d_EF,#d_Tipo,#d_Tam,#d_IMO,#d_Navio,#d_Sigvig').val('');
        $('#tblBL tbody').empty();
    }

    function preencherDetalhe(d) {
        state.detalhe = d;
        $('#txtIdCompleto').val(d.ID_CONTEINER || '');
        $('#txtYardAtual').val(d.YARD_ATUAL || '');
        $('#txtTara').val(d.TARA != null ? d.TARA : '');
        $('#txtLacreIpa').val(d.LACRE_IPA || '');
        $('#txtLacre1').val(d.LACRE1_GATE || '');
        $('#txtLacre2').val(d.LACRE2_GATE || '');
        $('#txtLacre3').val(d.LACRE3_GATE || '');
        $('#txtLacreLona').val(d.LACRE_LONA || '');
        $('#txtContexto').val(d.TEXTO_CONTEXTUAL || '');

        var badges = [];
        if (d.LBL_DTA_VISIVEL) badges.push('<span class="label label-danger">DTA</span>');
        if (d.LBL_DESOVA_VISIVEL) badges.push('<span class="label label-warning">' + (d.LBL_DESOVA_TEXTO || '') + '</span>');
        if (d.FLAG_REEFER_DESLIGADO_VISIVEL) badges.push('<span class="label label-danger">REEFER DESLIGADO</span>');
        if (d.FLAG_REEFER_LIGADO_VISIVEL) badges.push('<span class="label label-primary">REEFER LIGADO</span>');
        if (d.ALERTA_LINHA_9_VISIVEL && d.ALERTA_LINHA_9) badges.push('<span class="label label-info">' + d.ALERTA_LINHA_9 + '</span>');
        $('#badges').html(badges.join(' '));

        $('#d_EF').val(d.EF || '');
        $('#d_Tipo').val(d.TIPO_BASICO || '');
        $('#d_Tam').val(d.TAMANHO != null ? d.TAMANHO : '');
        $('#d_IMO').val(d.IMO || '');
        $('#d_Navio').val(d.NAVIO_VIAGEM || '');
        $('#d_Sigvig').val(d.SIGVIG || '');
        var tb = $('#tblBL tbody');
        tb.empty();
        if (d.LOTES_BL && d.LOTES_BL.length) {
            d.LOTES_BL.forEach(function (x) {
                tb.append('<tr><td>' + x.LOTE + '</td><td>' + (x.BL || '') + '</td></tr>');
            });
        }

        if (d.AVISO_REEFER_POSICAO_MSG)
            alerta(d.AVISO_REEFER_POSICAO_MSG, d.AVISO_REEFER_POSICAO_BLOQUEADO ? 'danger' : 'info');

        carregarProximos(d.ID_CONTEINER);
        carregarVizinhos();
        carregarHistoricosUi();
        if (mcpPerm.avarias) carregarAvariasUi();
    }

    function carregarProximos(id) {
        $.post(mcpUrls.proximos, { idConteiner: id }).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblProximos tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + (x.MOTIVO || '') + '</td><td>' + (x.DT_PREV || '') + '</td></tr>');
            });
        });
    }

    function modoPilhaDestino() {
        return $('input[name="modoPilha"]:checked').val() === 'destino';
    }

    function carregarVizinhos() {
        if (!state.detalhe) return;
        var req = {
            ID_CONTEINER: state.detalhe.ID_CONTEINER,
            YARD_ATUAL: state.detalhe.YARD_ATUAL,
            PREFIXO_DESTINO: $('#txtDestino').val(),
            MODO_DESTINO: modoPilhaDestino()
        };
        $.post(mcpUrls.vizinhos, req).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblVizinhos tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + (x.ID_CONTEINER || '') + '</td><td>' + (x.YARD || '') + '</td><td>' + (x.DT_PREV || '') + '</td><td>' + (x.MOTIVO || '') + '</td></tr>');
            });
        });
    }

    function carregarHistoricosUi() {
        if (!state.detalhe) return;
        $.get(mcpUrls.histShift, { idConteiner: state.detalhe.ID_CONTEINER, sistema: state.detalhe.SISTEMA }).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblHistShift tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + (x.DATA || '') + '</td><td>' + (x.ORIGEM || '') + '</td><td>' + (x.DESTINO || '') + '</td><td>' + (x.NMUSUARIO || '') + '</td><td>' + (x.DESC_MOTIVO || '') + '</td></tr>');
            });
        });
        $.get(mcpUrls.avHist, { autonumCntr: state.detalhe.AUTONUM }).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblHistAv tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                tb.append('<tr><td>' + (x.DT_AVARIA || '') + '</td><td>' + (x.AVARIA || '') + '</td><td>' + (x.USUARIO || '') + '</td></tr>');
            });
        });
    }

    function carregarAvariasUi() {
        if (!state.detalhe || !mcpPerm.avarias) return;
        $.get(mcpUrls.avGrid, { autonumCntr: state.detalhe.AUTONUM }).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblAvAbertas tbody');
            tb.empty();
            (r.dados || []).forEach(function (x) {
                var cls = x.DESTACAR_OPERADOR ? ' class="warning"' : '';
                tb.append('<tr' + cls + ' data-autonum="' + x.AUTONUM + '"><td>' + (x.LOCAL || '') + '</td><td>' + (x.TIPO || '') + '</td><td>' + x.AUTONUM + '</td></tr>');
            });
        });
        $.get(mcpUrls.avOpcoes).done(function (r) {
            if (!r.success) return;
            var dl = $('#ddlAvLocal');
            var dt = $('#ddlAvTipo');
            dl.empty();
            dt.empty();
            (r.locais || []).forEach(function (x) {
                dl.append($('<option/>').val(x).text(x));
            });
            (r.tipos || []).forEach(function (x) {
                dt.append($('<option/>').val(x).text(x));
            });
        });
    }

    function buscarPorFinal() {
        var f = ($('#txtFinal').val() || '').trim();
        if (!f) return;
        $.post(mcpUrls.buscarFinal, { final: f }).done(function (r) {
            if (!r.success) {
                alerta(r.message, 'danger');
                return;
            }
            $('#txtIdCompleto').val(r.idConteiner);
            carregarDetalhe();
        });
    }

    function carregarDetalhe() {
        var id = ($('#txtIdCompleto').val() || '').trim();
        if (!id) {
            alerta('Informe o ID do conteiner', 'danger');
            return;
        }
        $.post(mcpUrls.detalhe, { idConteiner: id }).done(function (r) {
            if (!r.success) {
                alerta(r.message, 'danger');
                return;
            }
            preencherDetalhe(r.data);
            alerta('Conteiner carregado', 'success');
        });
    }

    function motivoExigeCamera() {
        var opt = $('#ddlMotivo option:selected');
        return opt.data('camera') === 1 || opt.attr('data-camera') === '1';
    }

    function montarPayloadConfirmar() {
        var d = state.detalhe;
        if (!d) return null;
        return {
            AUTONUM_CNTR: d.AUTONUM,
            ID_CONTEINER: d.ID_CONTEINER,
            SISTEMA: d.SISTEMA,
            YARD_ORIGEM: d.YARD_ATUAL,
            YARD_DESTINO: ($('#txtDestino').val() || '').trim().toUpperCase(),
            MOTIVO_ID: parseInt($('#ddlMotivo').val(), 10) || 0,
            CAMERA_ID: $('#ddlCamera').val() ? parseInt($('#ddlCamera').val(), 10) : null,
            TAMANHO: d.TAMANHO || 0,
            EF: d.EF || '',
            TIPO_BASICO: d.TIPO_BASICO || '',
            TARA: d.TARA || 0,
            GMCI: d.GMCI || '',
            LACRE_IPA: $('#txtLacreIpa').val(),
            LACRE1_GATE: $('#txtLacre1').val(),
            LACRE2_GATE: $('#txtLacre2').val(),
            LACRE3_GATE: $('#txtLacre3').val(),
            LACRE_LONA: $('#txtLacreLona').val(),
            LACRE_IPA_TAG_ORIGEM: d.LACRE_IPA_TAG_ORIGEM || '',
            LACRE1_TAG: d.LACRE1_TAG || '',
            LACRE2_TAG: d.LACRE2_TAG || '',
            LACRE3_TAG: d.LACRE3_TAG || '',
            LACRES_CONFERIDOS: $('#chkLacresConf').is(':checked'),
            FLAG_CNTR_BLOQ_CHAO_PATIO: d.FLAG_CNTR_BLOQ_CHAO_PATIO || 0,
            IMO: d.IMO || '',
            CONFIRMOU_SAIDA: state.usuarioConfirmouSaida,
            CONFIRMOU_VEICULO_SERV: state.usuarioConfirmouVeiculoServ
        };
    }

    function enviarConfirmacao() {
        if (state.salvando) return;
        var payload = montarPayloadConfirmar();
        if (!payload) {
            alerta('Carregue um conteiner antes', 'danger');
            return;
        }

        var mid = parseInt($('#ddlMotivo').val(), 10);
        if (!mid) {
            alerta('Selecione o motivo da movimentacao', 'danger');
            return;
        }
        var dest = ($('#txtDestino').val() || '').trim();
        if (!dest) {
            alerta('Informe o yard de destino', 'danger');
            return;
        }

        if (motivoExigeCamera()) {
            var cam = $('#ddlCamera').val();
            if (!cam) {
                state.pendenteCamera = true;
                $('#modalCamera').modal('show');
                alerta('Selecione a camera', 'info');
                return;
            }
        }

        state.salvando = true;
        $.post(mcpUrls.confirmar, payload).done(function (r) {
            state.salvando = false;
            if (r.exigeCamera) {
                $('#modalCamera').modal('show');
                alerta(r.message, 'info');
                return;
            }
            if (r.exigeConfirmacaoSaida) {
                if (confirm('Confirma Saida da Unidade ?')) {
                    state.usuarioConfirmouSaida = true;
                    enviarConfirmacao();
                } else {
                    alerta('Operacao cancelada', 'info');
                }
                return;
            }
            if (r.exigeVeiculoServ) {
                if (confirm(r.message + '\n\nConfirma prosseguir (equivalente ao modal Veiculo/Servico legado)?')) {
                    state.usuarioConfirmouVeiculoServ = true;
                    enviarConfirmacao();
                }
                return;
            }
            if (!r.success) {
                alerta(r.message, 'danger');
                return;
            }
            alerta(r.message, 'success');
            limparFormMov();
        }).fail(function () {
            state.salvando = false;
            alerta('Erro na requisicao', 'danger');
        });
    }

    $('#btnBuscar').click(function () {
        var id = ($('#txtIdCompleto').val() || '').trim();
        var fin = ($('#txtFinal').val() || '').trim();
        if (id) {
            carregarDetalhe();
        } else if (fin) {
            buscarPorFinal();
        } else {
            alerta('Informe o final ou o ID completo do conteiner', 'danger');
        }
    });

    $('#btnLimpar').click(function () {
        limparFormMov();
        $('#alert-container').empty();
    });

    $('#txtDestino').on('keyup change', function () {
        this.value = this.value.toUpperCase();
        carregarVizinhos();
        var y = $(this).val().trim();
        if (y.length > 0) {
            $.post(mcpUrls.cameraPadrao, { yardDestino: y }).done(function (r) {
                if (r.success && r.cameraId != null)
                    $('#ddlCamera').val(String(r.cameraId));
            });
        }
    });

    $('input[name="modoPilha"]').change(carregarVizinhos);

    $('#btnSalvar').click(function () {
        state.usuarioConfirmouSaida = false;
        state.usuarioConfirmouVeiculoServ = false;
        enviarConfirmacao();
    });

    $('#btnConfirmarCamera').click(function () {
        $('#modalCamera').modal('hide');
        enviarConfirmacao();
    });

    $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
        var target = $(e.target).attr('href');
        if (target === '#tab-hist' || target === '#tab-dados')
            carregarHistoricosUi();
        if (target === '#tab-av')
            carregarAvariasUi();
    });

    $('#tblAvAbertas').on('click', 'tr', function () {
        $('#tblAvAbertas tr').removeClass('info');
        $(this).addClass('info');
    });

    $('#btnAvIncluir').click(function () {
        if (!state.detalhe) return;
        var loc = $('#ddlAvLocal').val();
        var tipos = $('#ddlAvTipo').val() || [];
        if (!loc) {
            alerta('Selecione o local', 'danger');
            return;
        }
        if (!tipos.length) {
            alerta('Selecione ao menos um tipo', 'danger');
            return;
        }
        $.ajax({
            url: mcpUrls.avIncluir,
            type: 'POST',
            traditional: true,
            data: {
                AUTONUM_CNTR: state.detalhe.AUTONUM,
                LOCAL_DESCR: loc,
                TIPOS_DESCR: tipos
            }
        }).done(function (r) {
            alerta(r.message, r.success ? 'success' : 'danger');
            if (r.success) carregarAvariasUi();
        }).fail(function () {
            alerta('Erro ao incluir avaria', 'danger');
        });
    });

    $('#btnAvExcluir').click(function () {
        if (!state.detalhe) return;
        var row = $('#tblAvAbertas tr.info');
        if (!row.length) {
            alerta('Selecione uma linha', 'danger');
            return;
        }
        var aut = parseInt(row.data('autonum'), 10);
        if (!confirm('Deseja excluir a avaria selecionada?')) return;
        $.post(mcpUrls.avExcluir, { AUTONUM_CNTR: state.detalhe.AUTONUM, AUTONUM_AVARIA: aut }).done(function (r) {
            alerta(r.message, r.success ? 'success' : 'danger');
            if (r.success) carregarAvariasUi();
        });
    });

    $('#btnAvFinalizar').click(function () {
        if (!state.detalhe) return;
        if (!confirm('Confirma a finalizacao das avarias do conteiner ' + state.detalhe.ID_CONTEINER + '?')) return;
        $.post(mcpUrls.avFinalizar, { AUTONUM_CNTR: state.detalhe.AUTONUM, ID_CONTEINER: state.detalhe.ID_CONTEINER }).done(function (r) {
            alerta(r.message, r.success ? 'success' : 'danger');
            if (r.success) carregarAvariasUi();
        });
    });
})();
