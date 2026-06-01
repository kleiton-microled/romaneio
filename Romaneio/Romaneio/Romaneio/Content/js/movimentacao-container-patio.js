(function () {
    var state = {
        detalhe: null,
        pendenteCamera: false,
        usuarioConfirmouSaida: false,
        usuarioConfirmouVeiculoServ: false,
        salvando: false,
        buscaFinalTimer: null,
        buscaXhr: null,
        avariasFinalizadas: false
    };

    function alerta(msg, tipo) {
        tipo = tipo || 'info';
        var cls = tipo === 'danger' ? 'alert-danger' : tipo === 'success' ? 'alert-success' : 'alert-info';
        var html = '<div class="alert ' + cls + ' alert-dismissible fade in"><button type="button" class="close" data-dismiss="alert">&times;</button>' + (msg || '') + '</div>';
        $('#alert-container').html(html);
    }

    function swalOk() {
        return window.Swal && typeof window.Swal.fire === 'function';
    }

    /** Normaliza JSON da API (camelCase ou PascalCase do MVC). */
    function normalizarRespostaApi(r) {
        if (r == null) {
            return { success: false, message: '', data: null, idConteiner: '' };
        }
        if (typeof r === 'string') {
            try {
                r = JSON.parse(r);
            } catch (e) {
                return { success: false, message: r, data: null, idConteiner: '' };
            }
        }
        var success = r.success !== undefined ? r.success : r.Success;
        if (typeof success === 'string') {
            success = success.toLowerCase() === 'true';
        }
        return {
            success: !!success,
            message: (r.message != null ? r.message : r.Message) || '',
            data: r.data != null ? r.data : r.Data,
            idConteiner: r.idConteiner || r.IdConteiner || r.ID_CONTEINER || ''
        };
    }

    function parseRespostaAjax(data, xhr) {
        if (data != null && typeof data === 'object') {
            return normalizarRespostaApi(data);
        }
        var texto = (typeof data === 'string' ? data : '') || (xhr && xhr.responseText) || '';
        if (!texto) {
            return { success: false, message: '', data: null, idConteiner: '' };
        }
        return normalizarRespostaApi(texto);
    }

    function abortarBuscaPendente() {
        if (state.buscaXhr && state.buscaXhr.readyState !== 4) {
            state.buscaXhr.abort();
        }
        state.buscaXhr = null;
    }

    /** Modal SweetAlert com a message retornada pela API (sem loading que bloqueia o 2o fire). */
    function exibirSwalBusca(norm, titulo, opcoes) {
        opcoes = opcoes || {};
        titulo = titulo || 'Busca';
        var ok = norm.success;
        var msg = norm.message ? String(norm.message) : '';
        if (ok && !msg) {
            msg = 'Conteiner carregado com sucesso.';
        }
        if (!ok && !msg) {
            msg = 'Nao foi possivel concluir a operacao.';
        }
        if (ok && opcoes.somenteErro) {
            return ok;
        }

        if (swalOk()) {
            window.Swal.fire({
                icon: ok ? 'success' : 'error',
                title: titulo,
                text: msg,
                confirmButtonText: 'OK',
                allowOutsideClick: true,
                heightAuto: false
            });
        } else {
            window.alert(titulo + '\n\n' + msg);
            alerta(msg, ok ? 'success' : 'danger');
        }
        return ok;
    }

    function postBuscaJson(url, data) {
        abortarBuscaPendente();
        state.buscaXhr = $.ajax({
            url: url,
            type: 'POST',
            data: data,
            dataType: 'text',
            cache: false
        });
        return state.buscaXhr.then(
            function (text, status, xhr) {
                state.buscaXhr = null;
                return parseRespostaAjax(text, xhr);
            },
            function (xhr, status) {
                state.buscaXhr = null;
                if (status === 'abort') {
                    return $.Deferred().reject('abort').promise();
                }
                var norm = parseRespostaAjax(null, xhr);
                if (!norm.message) {
                    norm.message = 'Erro ao comunicar com o servidor';
                }
                norm.success = false;
                return $.Deferred().reject(norm).promise();
            }
        );
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
        $('#tblVizinhos tbody').html('<tr class="text-muted"><td colspan="4">Carregue um conteiner</td></tr>');
        $('#alertTransporteInterno').hide();
        limparAbaDados();
        limparAbaHistorico();
        limparAbaAvarias();
    }

    function limparAbaHistorico() {
        $('#histPlaceholder').show();
        $('#histConteudo').hide();
        $('#histCntrTitulo').text('');
        $('#tblHistShift tbody').empty();
        $('#tblHistAv tbody').empty();
    }

    function atualizarVisibilidadeHistorico() {
        if (!state.detalhe) {
            limparAbaHistorico();
            return;
        }
        $('#histPlaceholder').hide();
        $('#histConteudo').show();
        $('#histCntrTitulo').text('Conteiner: ' + (state.detalhe.ID_CONTEINER || '') + ' (sistema ' + (state.detalhe.SISTEMA || '') + ')');
    }

    function atualizarLabelPilha() {
        var lbl = modoPilhaDestino() ? 'Q/FIADA DESTINO' : 'FIADA ATUAL';
        $('#lblPilhaRef').text(lbl);
    }

    /** Legado: mskDestino_LostFocus / TXTYardAtual_LostFocus (motivos 14, 1, 8). */
    function aplicarMotivoSugerido() {
        if (!state.detalhe) return;
        var yard = ($('#txtYardAtual').val() || '').trim().toUpperCase();
        var dest = ($('#txtDestino').val() || '').trim().toUpperCase();
        var id = ($('#txtIdCompleto').val() || '').trim().toUpperCase();
        if (!yard || !id) return;

        var motivoId = null;
        if (!yard || yard === 'GATEIN') {
            if (dest !== 'SAIDA') motivoId = 14;
        }
        if (yard && id !== 'SAIDA') {
            motivoId = 1;
        }
        if (dest === 'SAIDA') {
            motivoId = 8;
        }
        if (motivoId != null && $('#ddlMotivo option[value="' + motivoId + '"]').length) {
            $('#ddlMotivo').val(String(motivoId));
        }
    }

    function limparAbaDados() {
        $('#dadosPlaceholder').show();
        $('#dadosConteudo').hide();
        $('#d_IdCntr').text('');
        $('#d_SistemaBadge').text('');
        $('#d_EF,#d_Tipo,#d_Tam,#d_IMO,#d_Navio,#d_Sigvig,#d_GateIn,#d_Yard,#d_POD,#d_Bruto').val('');
        $('#d_ISO,#d_Tara,#d_MaxGross,#d_GMCI,#d_LacresManifesto,#d_TempSet,#d_TempMax,#d_TempMin').val('');
        $('#d_ReeferRow').hide();
        $('#tblBL tbody').empty();
    }

    /** Legado: campos da aba Dados CNTR preenchidos em Busca_Dados + DataGrid2 (BLs). */
    function preencherAbaDados(d) {
        if (!d) {
            limparAbaDados();
            return;
        }
        $('#dadosPlaceholder').hide();
        $('#dadosConteudo').show();
        $('#d_IdCntr').text(d.ID_CONTEINER || '');
        $('#d_SistemaBadge').text(d.SISTEMA ? 'Sistema ' + d.SISTEMA : '');

        $('#d_EF').val(d.EF || '');
        $('#d_Tipo').val(d.TIPO_BASICO || '');
        $('#d_Tam').val(d.TAMANHO != null && d.TAMANHO !== '' ? d.TAMANHO : '');
        $('#d_IMO').val(d.IMO || '');
        $('#d_GateIn').val(d.TXT_GATE_IN || '');
        $('#d_Yard').val(d.YARD_ATUAL || '');
        $('#d_Navio').val(d.NAVIO_VIAGEM || '');
        $('#d_POD').val(d.POD || '');
        $('#d_Bruto').val(d.BRUTO || '');
        $('#d_Sigvig').val(d.SIGVIG || '');
        $('#d_ISO').val(d.ISO || '');
        $('#d_Tara').val(d.TARA != null ? d.TARA : '');
        $('#d_MaxGross').val(d.MAX_GROSS != null ? d.MAX_GROSS : '');
        $('#d_GMCI').val(d.GMCI || '');
        $('#d_LacresManifesto').val(d.LACRES_RESUMO || '');

        var exibeReefer = d.FLAG_REEFER_DESLIGADO_VISIVEL || d.FLAG_REEFER_LIGADO_VISIVEL ||
            (d.TEMP_SETPOINT && d.TEMP_SETPOINT.trim()) ||
            (d.TEMP_MAX && d.TEMP_MAX.trim()) ||
            (d.TEMP_MIN && d.TEMP_MIN.trim());
        if (exibeReefer) {
            $('#d_ReeferRow').show();
            $('#d_TempSet').val(d.TEMP_SETPOINT || '');
            $('#d_TempMax').val(d.TEMP_MAX || '');
            $('#d_TempMin').val(d.TEMP_MIN || '');
        } else {
            $('#d_ReeferRow').hide();
        }

        var tb = $('#tblBL tbody');
        tb.empty();
        var bls = d.LOTES_BL || [];
        if (!bls.length) {
            tb.append('<tr class="text-muted"><td colspan="2">Nenhum BL vinculado</td></tr>');
            return;
        }
        bls.forEach(function (x) {
            tb.append('<tr><td>' + x.LOTE + '</td><td>' + (x.BL || '') + '</td></tr>');
        });
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

        preencherAbaDados(d);

        if (d.FRAME_TRANSPORTE_INTERNO) {
            $('#alertTransporteInterno').show();
        } else {
            $('#alertTransporteInterno').hide();
        }

        if (d.AVISO_REEFER_POSICAO_MSG) {
            alerta(d.AVISO_REEFER_POSICAO_MSG, d.AVISO_REEFER_POSICAO_BLOQUEADO ? 'danger' : 'info');
        }

        carregarProximos(d.ID_CONTEINER);
        carregarVizinhos();
        aplicarMotivoSugerido();
        if ($('#tab-hist').hasClass('active')) {
            carregarAbaHistorico();
        }
        if ($('#tab-dados').hasClass('active')) {
            preencherAbaDados(d);
        }
        if (mcpPerm.avarias && $('#tab-av').hasClass('active')) {
            carregarAbaAvarias();
        }
    }

    function carregarProximos(id) {
        $.post(mcpUrls.proximos, { idConteiner: id }).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblProximos tbody');
            tb.empty();
            var rows = r.dados || [];
            if (!rows.length) {
                tb.append('<tr class="text-muted"><td colspan="2">Nenhum movimento previsto</td></tr>');
                return;
            }
            rows.forEach(function (x) {
                tb.append('<tr><td>' + (x.MOTIVO || '') + '</td><td>' + (x.DT_PREV || '') + '</td></tr>');
            });
        }).fail(function () {
            alerta('Erro ao carregar proximos movimentos', 'danger');
        });
    }

    function modoPilhaDestino() {
        return $('input[name="modoPilha"]:checked').val() === 'destino';
    }

    function carregarVizinhos() {
        if (!state.detalhe) return;
        var dest = ($('#txtDestino').val() || '').trim();
        if (!dest && modoPilhaDestino()) {
            $('#tblVizinhos tbody').empty();
            return;
        }
        var req = {
            ID_CONTEINER: state.detalhe.ID_CONTEINER,
            YARD_ATUAL: state.detalhe.YARD_ATUAL,
            PREFIXO_DESTINO: dest,
            MODO_DESTINO: modoPilhaDestino()
        };
        $.post(mcpUrls.vizinhos, req).done(function (r) {
            if (!r.success) return;
            var tb = $('#tblVizinhos tbody');
            tb.empty();
            var rows = r.dados || [];
            if (!rows.length) {
                tb.append('<tr class="text-muted"><td colspan="4">Nenhuma unidade na pilha</td></tr>');
                return;
            }
            rows.forEach(function (x) {
                tb.append('<tr><td>' + (x.ID_CONTEINER || '') + '</td><td>' + (x.YARD || '') + '</td><td>' + (x.DT_PREV || '') + '</td><td>' + (x.MOTIVO || '') + '</td></tr>');
            });
        }).fail(function () {
            alerta('Erro ao carregar pilha / vizinhos', 'danger');
        });
    }

  /** Legado: SSTab1 tab 1 — Carrega_Hist_Avaria + SQL shifting por sistema (I/A/R). */
    function carregarAbaHistorico() {
        if (!state.detalhe) {
            limparAbaHistorico();
            return;
        }
        atualizarVisibilidadeHistorico();

        var tbShift = $('#tblHistShift tbody');
        var tbAv = $('#tblHistAv tbody');
        tbShift.html('<tr class="text-muted"><td colspan="7">Carregando...</td></tr>');
        tbAv.html('<tr class="text-muted"><td colspan="3">Carregando...</td></tr>');

        $.get(mcpUrls.histShift, {
            idConteiner: state.detalhe.ID_CONTEINER,
            sistema: state.detalhe.SISTEMA
        }).done(function (r) {
            tbShift.empty();
            if (!r.success) {
                alerta(r.message || 'Erro ao carregar historico de shifting', 'danger');
                tbShift.append('<tr class="text-muted"><td colspan="7">Falha na consulta</td></tr>');
                return;
            }
            var rows = r.dados || [];
            if (!rows.length) {
                tbShift.append('<tr class="text-muted"><td colspan="7">Nenhuma movimentacao registrada</td></tr>');
                return;
            }
            rows.forEach(function (x) {
                tbShift.append(
                    '<tr><td>' + (x.DATA || '') + '</td><td>' + (x.ORIGEM || '') + '</td><td>' + (x.DESTINO || '') + '</td>' +
                    '<td>' + (x.NMUSUARIO || '') + '</td><td>' + (x.IDENTIFICACAO || '') + '</td><td>' + (x.DESC_MOTIVO || '') + '</td>' +
                    '<td>' + (x.VEICULO || '') + '</td></tr>'
                );
            });
        }).fail(function () {
            tbShift.html('<tr class="text-muted"><td colspan="7">Erro ao comunicar com o servidor</td></tr>');
            alerta('Erro ao carregar historico de shifting', 'danger');
        });

        $.get(mcpUrls.avHist, { autonumCntr: state.detalhe.AUTONUM }).done(function (r) {
            tbAv.empty();
            if (!r.success) {
                alerta(r.message || 'Erro ao carregar historico de avarias', 'danger');
                tbAv.append('<tr class="text-muted"><td colspan="3">Falha na consulta</td></tr>');
                return;
            }
            var rows = r.dados || [];
            if (!rows.length) {
                tbAv.append('<tr class="text-muted"><td colspan="3">Nenhuma avaria registrada</td></tr>');
                return;
            }
            rows.forEach(function (x) {
                tbAv.append('<tr><td>' + (x.DT_AVARIA || '') + '</td><td>' + (x.AVARIA || '') + '</td><td>' + (x.USUARIO || '') + '</td></tr>');
            });
        }).fail(function () {
            tbAv.html('<tr class="text-muted"><td colspan="3">Erro ao comunicar com o servidor</td></tr>');
            alerta('Erro ao carregar historico de avarias', 'danger');
        });
    }

    function limparAbaAvarias() {
        if (!mcpPerm.avarias) return;
        state.avariasFinalizadas = false;
        $('#avPlaceholder').show();
        $('#avConteudo').hide();
        $('#avCntrTitulo').text('');
        $('#avAlertaFinalizada').hide();
        $('#tblAvAbertas tbody').empty();
        $('#ddlAvLocal').html('<option value="">Selecione</option>');
        $('#ddlAvTipo').empty();
        setaControlesAvarias();
    }

    /** Legado: SetaControlesAV — desabilita acoes se patio ja finalizado. */
    function setaControlesAvarias() {
        if (!mcpPerm.avarias) return;
        var bloqueado = !state.detalhe || state.avariasFinalizadas;
        $('#btnAvIncluir').prop('disabled', bloqueado || !mcpPerm.avIncluir);
        $('#btnAvExcluir').prop('disabled', bloqueado || !mcpPerm.avExcluir);
        $('#btnAvFinalizar').prop('disabled', bloqueado || !mcpPerm.avFinalizar);
        $('#ddlAvLocal, #ddlAvTipo').prop('disabled', bloqueado);
    }

    /** Legado: SSTab1 tab 3 — Carrega_ListLocal, Carrega_ListTipo, CarregaAvarias. */
    function carregarAbaAvarias() {
        if (!mcpPerm.avarias) return;
        if (!state.detalhe) {
            limparAbaAvarias();
            return;
        }

        $('#avPlaceholder').hide();
        $('#avConteudo').show();
        $('#avCntrTitulo').text('Conteiner: ' + (state.detalhe.ID_CONTEINER || ''));

        var tb = $('#tblAvAbertas tbody');
        tb.html('<tr class="text-muted"><td colspan="2">Carregando...</td></tr>');

        $.get(mcpUrls.avOpcoes).done(function (r) {
            if (!r.success) return;
            var dl = $('#ddlAvLocal');
            var dt = $('#ddlAvTipo');
            dl.html('<option value="">Selecione</option>');
            dt.empty();
            (r.locais || []).forEach(function (x) {
                dl.append($('<option/>').val(x).text(x));
            });
            (r.tipos || []).forEach(function (x) {
                dt.append($('<option/>').val(x).text(x));
            });
            setaControlesAvarias();
        }).fail(function () {
            alerta('Erro ao carregar opcoes de avaria', 'danger');
        });

        $.get(mcpUrls.avGrid, { autonumCntr: state.detalhe.AUTONUM }).done(function (r) {
            tb.empty();
            if (!r.success) {
                alerta(r.message || 'Erro ao carregar avarias', 'danger');
                tb.append('<tr class="text-muted"><td colspan="2">Falha na consulta</td></tr>');
                return;
            }

            state.avariasFinalizadas = !!r.finalizada;
            if (state.avariasFinalizadas) {
                $('#avAlertaFinalizada').show();
            } else {
                $('#avAlertaFinalizada').hide();
            }
            setaControlesAvarias();

            var rows = r.dados || [];
            if (!rows.length) {
                tb.append('<tr class="text-muted"><td colspan="2">Nenhuma avaria aberta</td></tr>');
                return;
            }
            rows.forEach(function (x) {
                var cls = x.DESTACAR_OPERADOR ? ' class="warning"' : '';
                var title = x.DESTACAR_OPERADOR ? ' title="Avaria apontada pelo operador"' : '';
                tb.append(
                    '<tr' + cls + title + ' data-autonum="' + x.AUTONUM + '">' +
                    '<td>' + (x.LOCAL || '') + '</td><td>' + (x.TIPO || '') + '</td></tr>'
                );
            });
        }).fail(function () {
            tb.html('<tr class="text-muted"><td colspan="2">Erro ao comunicar com o servidor</td></tr>');
            alerta('Erro ao carregar avarias abertas', 'danger');
        });
    }

    function carregarDetalheInterno(opcoes) {
        opcoes = opcoes || {};
        var id = ($('#txtIdCompleto').val() || '').trim();
        return postBuscaJson(mcpUrls.detalhe, { idConteiner: id }).then(function (norm) {
            if (!norm.success) {
                exibirSwalBusca(norm, 'Conteiner');
                return norm;
            }
            preencherDetalhe(norm.data);
            if (!opcoes.silenciarSucesso) {
                exibirSwalBusca(norm, 'Conteiner', { somenteErro: false });
            }
            return norm;
        }).fail(function (err) {
            if (err === 'abort') return;
            var norm = (err && err.message !== undefined) ? err : { success: false, message: 'Erro ao comunicar com o servidor' };
            exibirSwalBusca(norm, 'Conteiner');
        });
    }

    function buscarPorFinal() {
        var f = ($('#txtFinal').val() || '').trim();
        if (!f) return;
        if (state.buscaFinalTimer) {
            clearTimeout(state.buscaFinalTimer);
            state.buscaFinalTimer = null;
        }
        postBuscaJson(mcpUrls.buscarFinal, { final: f }).then(function (norm) {
            if (!norm.success) {
                exibirSwalBusca(norm, 'Busca por final');
                return;
            }
            $('#txtIdCompleto').val(norm.idConteiner);
            carregarDetalheInterno({ silenciarSucesso: false });
        }).fail(function (err) {
            if (err === 'abort') return;
            var norm = (err && err.message !== undefined) ? err : { success: false, message: 'Erro ao comunicar com o servidor' };
            exibirSwalBusca(norm, 'Busca por final');
        });
    }

    function carregarDetalhe() {
        var id = ($('#txtIdCompleto').val() || '').trim();
        if (!id) {
            exibirSwalBusca({ success: false, message: 'Informe o ID do conteiner' }, 'Conteiner');
            return;
        }
        if (state.buscaFinalTimer) {
            clearTimeout(state.buscaFinalTimer);
            state.buscaFinalTimer = null;
        }
        carregarDetalheInterno({ silenciarSucesso: false });
    }

    function agendarBuscaPorFinal() {
        if (state.buscaFinalTimer) clearTimeout(state.buscaFinalTimer);
        var fin = ($('#txtFinal').val() || '').trim();
        if (!fin || fin.indexOf('_') >= 0) return;
        if (($('#txtIdCompleto').val() || '').trim()) return;
        state.buscaFinalTimer = setTimeout(function () {
            state.buscaFinalTimer = null;
            if (fin.length >= 4) buscarPorFinal();
        }, 450);
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

    $('#txtFinal').on('input', agendarBuscaPorFinal);

    $('#txtFinal, #txtIdCompleto').on('keypress', function (e) {
        if (e.which === 13) {
            e.preventDefault();
            $('#btnBuscar').click();
        }
    });

    $('#txtDestino').on('keyup change blur', function () {
        this.value = this.value.toUpperCase();
        aplicarMotivoSugerido();
        carregarVizinhos();
        var y = $(this).val().trim();
        if (y.length > 0) {
            $.post(mcpUrls.cameraPadrao, { yardDestino: y }).done(function (r) {
                if (r.success && r.cameraId != null) {
                    $('#ddlCamera').val(String(r.cameraId));
                }
            });
        }
    });

    $('input[name="modoPilha"]').change(function () {
        atualizarLabelPilha();
        carregarVizinhos();
    });

    atualizarLabelPilha();

    $('#btnSalvar').click(function () {
        state.usuarioConfirmouSaida = false;
        state.usuarioConfirmouVeiculoServ = false;
        enviarConfirmacao();
    });

    $('#btnConfirmarCamera').click(function () {
        $('#modalCamera').modal('hide');
        enviarConfirmacao();
    });

    $('a[data-toggle="tab"]').on('show.bs.tab', function (e) {
        var next = $(e.target).attr('href');
        if (next !== '#tab-mov' && !state.detalhe) {
            e.preventDefault();
            alerta('Carregue um conteiner na aba Movimentacao antes de consultar outras abas', 'info');
        }
    });

    $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
        var target = $(e.target).attr('href');
        if (target === '#tab-hist') carregarAbaHistorico();
        if (target === '#tab-dados' && state.detalhe) preencherAbaDados(state.detalhe);
        if (target === '#tab-av') carregarAbaAvarias();
    });

    /** Bootstrap 4: painel inicial precisa de .show (nao .in) para ficar visivel sem clicar na aba. */
    function garantirAbaMovimentacaoVisivel() {
        var $pane = $('#tab-mov');
        if (!$pane.length) return;
        $pane.addClass('active show');
        $('a[href="#tab-mov"]').parent('li').addClass('active');
    }

    $(garantirAbaMovimentacaoVisivel);

    $('#tblAvAbertas').on('click', 'tbody tr[data-autonum]', function () {
        $('#tblAvAbertas tbody tr').removeClass('info');
        $(this).addClass('info');
    });

    $('#btnAvIncluir').click(function () {
        if (!state.detalhe || state.avariasFinalizadas) return;
        var loc = $('#ddlAvLocal').val();
        var tipos = $('#ddlAvTipo').val() || [];
        if (!loc) {
            alerta('Nenhum Local selecionado', 'danger');
            return;
        }
        if (!tipos.length) {
            alerta('Nenhum tipo selecionado', 'danger');
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
            if (r.success) carregarAbaAvarias();
        }).fail(function () {
            alerta('Erro ao incluir avaria', 'danger');
        });
    });

    $('#btnAvExcluir').click(function () {
        if (!state.detalhe || state.avariasFinalizadas) return;
        var row = $('#tblAvAbertas tbody tr.info');
        if (!row.length) {
            alerta('Selecione uma linha na grid', 'danger');
            return;
        }
        var aut = parseInt(row.data('autonum'), 10);
        if (!aut) {
            alerta('Avaria apontada pelo Operador', 'danger');
            return;
        }
        if (!confirm('Deseja excluir a avaria selecionada?')) return;
        $.post(mcpUrls.avExcluir, { AUTONUM_CNTR: state.detalhe.AUTONUM, AUTONUM_AVARIA: aut }).done(function (r) {
            alerta(r.message, r.success ? 'success' : 'danger');
            if (r.success) carregarAbaAvarias();
        }).fail(function () {
            alerta('Erro ao excluir avaria', 'danger');
        });
    });

    $('#btnAvFinalizar').click(function () {
        if (!state.detalhe || state.avariasFinalizadas) return;
        if (!confirm('Confirma a finalizacao das avarias do conteiner ' + state.detalhe.ID_CONTEINER + '?')) return;
        $.post(mcpUrls.avFinalizar, { AUTONUM_CNTR: state.detalhe.AUTONUM, ID_CONTEINER: state.detalhe.ID_CONTEINER }).done(function (r) {
            alerta(r.message, r.success ? 'success' : 'danger');
            if (r.success) carregarAbaAvarias();
        }).fail(function () {
            alerta('Erro ao finalizar avarias', 'danger');
        });
    });
})();
