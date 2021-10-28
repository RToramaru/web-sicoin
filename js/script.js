$('#modalUsuario').on('show.bs.modal', function (event) {
    var button = $(event.relatedTarget)
    var nome = button.data('nome')
    var endereco = button.data('endereco')
    var telefone = button.data('telefone')
    var usuario = button.data('usuario')
    var tipo = button.data('tipo')
    var modal = $(this)

    modal.find('.modal-title').text('Editar usuário ' + nome)
    modal.find('#nome').val(nome)
    modal.find('#endereco').val(endereco)
    modal.find('#telefone').val(telefone)
    modal.find('#usuario').val(usuario)
    modal.find('#tipo').val(tipo)
})


$('#modalSetor').on('show.bs.modal', function (event) {
    var button = $(event.relatedTarget)
    var setor = button.data('setor')
    var responsavel = button.data('responsavel')
    var tipo = button.data('tipo')
    var cnpj = button.data('cnpj')
    var modal = $(this)

    modal.find('.modal-title').text('Editar setor ' + setor)
    modal.find('#setor').val(setor)
    modal.find('#responsavel').val(responsavel)
    modal.find('#tipo').val(tipo)
    modal.find('#cnpj').val(cnpj)
})



$('a[data-confirmar]').click(function (ev) {
    var href = $(this).attr('href')
    if (!$('#confirm-delete').length) {
        $('body').append('<div class="modal fade" id="confirm-delete" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true"><div class="modal-dialog"><div class="modal-content"><div class="modal-header bg-danger text-white">EXCLUIR ITEM<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button></div><div class="modal-body">Tem certeza de que deseja excluir o item selecionado?</div><div class="modal-footer"><button type="button" class="btn btn-success" data-dismiss="modal">Cancelar</button><a class="btn btn-danger text-white" id="dataComfirmOK">Apagar</a></div></div></div></div>')
    }
    $('#dataComfirmOK').attr('href', href)
    $('#confirm-delete').modal({ show: true })
    return false
})


