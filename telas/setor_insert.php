<?php

require "../banco_dados/conecta.php";

$setor = $_POST['setor'];
$responsavel = $_POST['responsavel'];
$tipo = $_POST['tipo'];
$cnpj = $_POST['cnpj'];

$sql = "INSERT INTO setor (setor, responsavel, tipo, cnpj)
VALUES ('$setor','$responsavel','$tipo', '$cnpj')";

if (mysqli_query($conn, $sql)) {
  ?>
<!DOCTYPE html>
<html>

<head>
    <meta charset='UTF-8' />
    <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css'>
</head>

<body>

    <div class='modal fade' id='myModal' tabindex='-1' role='dialog'>
        <div class='modal-dialog'>
            <div class='modal-content'>
                <div class="modal-header bg-success text-white">
                    <h4 class='modal-title'>Sucesso</h4>
                </div>
                <div class='modal-body'>
                    <h4>Cadastrado com sucesso!</h4>  
                </div>
                <div class='modal-footer'>
                    <input class='btn btn-success' data-dismiss='modal' value='Fechar' id='modal' onclick='confirmar()'>
                </div>
            </div>
        </div>
    </div>


    <script src='https://code.jquery.com/jquery-3.3.1.slim.min.js'></script>
    <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js'></script>

    <script>
    $(document).ready(function() {
        $('#myModal').modal('show')
    })

    function confirmar() {
        window.location.replace('setor.php')
    }
    </script>
</body>
<?php
} else {
    ?>
<!DOCTYPE html>
<html>

<head>
    <meta charset='UTF-8' />
    <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css'>
</head>

<body>

    <div class='modal fade' id='myModal' tabindex='-1' role='dialog'>
        <div class='modal-dialog'>
            <div class='modal-content'>
                <div class="modal-header bg-danger text-white">
                    <h4 class='modal-title'>Erro</h4>
                </div>
                <div class='modal-body'>
                    <h4>Ocorreu um erro ao cadastrar!</h4>  
                </div>
                <div class='modal-footer'>
                    <input class='btn btn-danger' data-dismiss='modal' value='Fechar' id='modal' onclick='confirmar()'>
                </div>
            </div>
        </div>
    </div>


    <script src='https://code.jquery.com/jquery-3.3.1.slim.min.js'></script>
    <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js'></script>

    <script>
    $(document).ready(function() {
        $('#myModal').modal('show')
    })

    function confirmar() {
        window.location.replace('setor.php')
    }
    </script>
</body>
<?php
}

?>