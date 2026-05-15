(function () {
    var state = {
        autonumRegistroSaida: 0,
        cabecalho: null,
        itens: [],
        operacaoLiberada: false
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

    function swalLoading(abrir) {
        if (!swalOk()) return;
        if (abrir) {
            window.Swal.fire({
                title: 'Consultando placa...',
                allowOutsideClick: false,
                didOpen: function () { window.Swal.showLoading(); }
            });
        } else if (typeof window.Swal.close === 'function') {
            window.Swal.close();
        }
    }

    function swalPlacaResposta(r) {
        var msg = (r && r.message) ? String(r.message) : '';
        if (!r || !r.success) {
            if (swalOk()) {
                window.Swal.fire({
                    icon: 'error',
                    title: 'Placa',
                    text: msg || 'Nao consta entrada para o veiculo informado',
                    confirmButtonText: 'OK'
                });
            } else {
                alerta(msg || 'Erro ao consultar placa', 'danger');
            }
            return false;
        }
        var lotes = r.lotes || r.LOTES || [];
        var texto = lotes.length > 0
            ? ('Entrada localizada. ' + lotes.length + ' lote(s) disponivel(is).')
            : 'Entrada localizada. Nenhum lote encontrado.';
        if (swalOk()) {
            window.Swal.fire({
                icon: 'success',
                title: 'Placa',
                text: texto,
                confirmButtonText: 'OK'
            });
        }
        return true;
    }

    function esc(s) {
        if (s == null || s === undefined) return '';
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function painelItensVisivel(v) {
        state.operacaoLiberada = !!v;
        $('#painelItens').toggle(!!v);
    }

    function limparFormItem() {
        $('#hidAutonumItem').val('0');
        $('#txtQuantidade').val('');
        $('#selEmbalagem').val('');
        $('input[name="radFinal"]').prop('checked', false);
    }

    function limparTudo() {
        state.autonumRegistroSaida = 0;
        state.cabecalho = null;
        state.itens = [];
        state.operacaoLiberada = false;
        $('#selPlaca').val('');
        $('#selLote').empty().append('<option value="">-- selecione --</option>').prop('disabled', true);
        $('#hidCntr,#hidGate,#hidLote,#hidIdOc').val('0');
        limparFormItem();
        $('#tblItens tbody').empty();
        painelItensVisivel(false);
    }

    function montarGrid() {
        var tb = $('#tblItens tbody');
        tb.empty();
        (state.itens || []).forEach(function (x) {
            var fin = (x.FINALIZADO === 1 || x.FINALIZADO === true || x.finalizado === 1) ? 'SIM' : 'NAO';
            tb.append(
                '<tr class="tr-item" data-autonum="' + esc(x.AUTONUM || x.autonum) + '" data-qtd="' + esc(x.QUANTIDADE || x.quantidade) + '" data-emb="' + esc(x.EMBALAGEM || x.embalagem) + '" data-fin="' + esc(x.FINALIZADO || x.finalizado) + '">' +
                '<td>' + esc(x.QUANTIDADE || x.quantidade) + '</td>' +
                '<td>' + esc(x.DESCR_EMB || x.descr_emb) + '</td>' +
                '<td>' + esc(fin) + '</td></tr>'
            );
        });
    }

    function carregarPlacas() {
        $.post(ddpUrls.placas).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro ao listar placas', 'danger');
                return;
            }
            var sel = $('#selPlaca');
            sel.find('option:not(:first)').remove();
            (r.dados || []).forEach(function (p) {
                var id = p.AUTONUM_RS || p.autonum_rs;
                var tx = p.PLACA || p.placa;
                sel.append('<option value="' + esc(id) + '">' + esc(tx) + '</option>');
            });
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function carregarEmbalagens() {
        $.post(ddpUrls.embalagens).done(function (r) {
            if (!r.success) return;
            var sel = $('#selEmbalagem');
            sel.empty().append('<option value="">-- selecione --</option>');
            (r.dados || []).forEach(function (e) {
                var c = e.CODE || e.code;
                var d = e.DESCR || e.descr;
                sel.append('<option value="' + esc(c) + '">' + esc(d) + '</option>');
            });
        });
    }

    function aoMudarPlaca() {
        var reg = parseInt($('#selPlaca').val(), 10) || 0;
        state.autonumRegistroSaida = reg;
        $('#selLote').empty().append('<option value="">-- selecione --</option>').prop('disabled', true);
        $('#hidCntr,#hidGate,#hidLote,#hidIdOc').val('0');
        limparFormItem();
        $('#tblItens tbody').empty();
        painelItensVisivel(false);
        if (reg <= 0) return;

        swalLoading(true);
        $.post(ddpUrls.gateLotes, { autonumRegistroSaida: reg }).done(function (r) {
            swalLoading(false);
            if (!swalPlacaResposta(r)) {
                $('#selPlaca').val('');
                state.autonumRegistroSaida = 0;
                return;
            }
            var lotes = r.lotes || r.LOTES || [];
            var sel = $('#selLote');
            sel.empty().append('<option value="">-- selecione --</option>');
            lotes.forEach(function (l) {
                var id = l.AUTONUM_LOTE || l.autonum_lote;
                sel.append('<option value="' + esc(id) + '">Lote ' + esc(id) + '</option>');
            });
            sel.prop('disabled', lotes.length === 0);
        }).fail(function () {
            swalLoading(false);
            $('#selPlaca').val('');
            state.autonumRegistroSaida = 0;
            if (swalOk()) {
                window.Swal.fire({
                    icon: 'error',
                    title: 'Placa',
                    text: 'Falha de rede ao consultar o veiculo.',
                    confirmButtonText: 'OK'
                });
            } else {
                alerta('Falha de rede', 'danger');
            }
        });
    }

    function carregarOperacao() {
        var reg = parseInt($('#selPlaca').val(), 10) || 0;
        var lote = parseInt($('#selLote').val(), 10) || 0;
        if (reg <= 0) {
            alerta('Placa nao informada', 'danger');
            return;
        }
        if (lote <= 0) {
            alerta('Lote nao informado', 'danger');
            return;
        }
        $.post(ddpUrls.carregar, { autonumRegistroSaida: reg, autonumLote: lote }).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro ao carregar', 'danger');
                return;
            }
            var cab = r.cabecalho || {};
            state.cabecalho = cab;
            state.itens = r.itens || [];
            $('#hidCntr').val(cab.CNTR || cab.cntr || 0);
            $('#hidGate').val(cab.GATE || cab.gate || 0);
            $('#hidLote').val(cab.LOTE || cab.lote || lote);
            $('#hidIdOc').val(cab.ID_OC || cab.id_oc || 0);
            montarGrid();
            painelItensVisivel(true);
            limparFormItem();
            alerta('Dados carregados.', 'success');
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function salvar() {
        if (!state.operacaoLiberada) {
            alerta('Nao disponivel', 'danger');
            return;
        }
        var reg = parseInt($('#selPlaca').val(), 10) || 0;
        var cntr = parseInt($('#hidCntr').val(), 10) || 0;
        var gate = parseInt($('#hidGate').val(), 10) || 0;
        var lote = parseInt($('#hidLote').val(), 10) || 0;
        var idOc = parseInt($('#hidIdOc').val(), 10) || 0;
        var qtd = parseFloat(String($('#txtQuantidade').val()).replace(',', '.')) || 0;
        var emb = ($('#selEmbalagem').val() || '').trim();
        var rad = $('input[name="radFinal"]:checked').val();
        if (rad === undefined) {
            alerta('Informe finalizacao de entrega', 'danger');
            return;
        }
        var fin = parseInt(rad, 10);
        var autonumItem = parseInt($('#hidAutonumItem').val(), 10) || 0;

        var payload = {
            AutonumRegistroSaida: reg,
            Cntr: cntr,
            Gate: gate,
            Lote: lote,
            IdOc: idOc,
            Quantidade: qtd,
            Embalagem: emb,
            Finalizado: fin,
            AutonumItem: autonumItem
        };

        $.post(ddpUrls.salvar, payload).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            var msg = r.message || 'OK';
            if (r.aviso) msg += ' ' + r.aviso;
            alerta(msg, 'success');
            var cntr = parseInt($('#hidCntr').val(), 10) || 0;
            var gate = parseInt($('#hidGate').val(), 10) || 0;
            $.post(ddpUrls.itens, { cntr: cntr, gate: gate }).done(function (r2) {
                if (r2.success) {
                    state.itens = r2.itens || [];
                    montarGrid();
                }
            });
            limparFormItem();
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    function novo() {
        if (!state.operacaoLiberada) {
            alerta('Nao disponivel', 'danger');
            return;
        }
        var reg = parseInt($('#selPlaca').val(), 10) || 0;
        var lote = parseInt($('#selLote').val(), 10) || 0;
        if (reg <= 0) {
            alerta('Placa nao informada', 'danger');
            return;
        }
        if (lote <= 0) {
            alerta('Lote nao informada', 'danger');
            return;
        }
        limparFormItem();
        $('#txtQuantidade').focus();
    }

    function excluir() {
        if (!state.operacaoLiberada) {
            alerta('Nao disponivel', 'danger');
            return;
        }
        var id = parseInt($('#hidAutonumItem').val(), 10) || 0;
        if (id <= 0) {
            alerta('Selecione um item na lista', 'danger');
            return;
        }
        if (!window.confirm('Confirma exclusao do item ?')) return;
        $.post(ddpUrls.excluir, { autonumItem: id }).done(function (r) {
            if (!r.success) {
                alerta(r.message || 'Erro', 'danger');
                return;
            }
            alerta(r.message || 'Excluido', 'success');
            var cntr = parseInt($('#hidCntr').val(), 10) || 0;
            var gate = parseInt($('#hidGate').val(), 10) || 0;
            limparFormItem();
            $.post(ddpUrls.itens, { cntr: cntr, gate: gate }).done(function (r2) {
                if (r2.success) {
                    state.itens = r2.itens || [];
                    montarGrid();
                }
            });
        }).fail(function () { alerta('Falha de rede', 'danger'); });
    }

    $(document).on('click', '#tblItens tbody tr.tr-item', function () {
        var tr = $(this);
        $('#hidAutonumItem').val(tr.data('autonum'));
        $('#txtQuantidade').val(tr.data('qtd'));
        $('#selEmbalagem').val(String(tr.data('emb')));
        var fin = parseInt(tr.data('fin'), 10) || 0;
        if (fin === 1) $('#radFinalSim').prop('checked', true);
        else $('#radFinalNao').prop('checked', true);
        $('#txtQuantidade').focus();
    });

    $(function () {
        carregarPlacas();
        carregarEmbalagens();

        $('#selPlaca').on('change', aoMudarPlaca);
        $('#btnCarregar').on('click', carregarOperacao);
        $('#btnLimpar').on('click', function () {
            limparTudo();
            carregarPlacas();
            alerta('Campos limpos.', 'info');
        });
        $('#btnNovo').on('click', novo);
        $('#btnSalvar').on('click', salvar);
        $('#btnExcluir').on('click', excluir);
    });
})();
