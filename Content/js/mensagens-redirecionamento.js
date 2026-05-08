var MensagemSucessoRedirecionar = function (resultado) {
    if (resultado.url) {
        console.log(resultado.url);
        window.location.href = resultado.url;
    }
};

var MensagemErro = function (xhr, status) {

    $('#btnPesquisar')
        .html('Pesquisar')
        .removeClass('disabled');
    if (xhr.responseText !== '') {
        var resultado = xhr.statusText;
        toastr.error(resultado, 'Portal');
        return;
    }
    if (xhr.statusText !== '') {
        var resultado = xhr.statusText;
        toastr.error(resultado, 'Portal');
        return;
    }
    toastr.error('Erro ao pesquisar item', 'Portal');
};