<?php

require "../banco_dados/conecta.php";



$codigo = $_POST['codigo'];
$origem = $_POST['origem'];
$destino = $_POST['destino'];
$emissao = $_POST['emissao'];
$recebimento = $_POST['recebimento'];
$dados = $_POST['dados'];


$dados_array = json_decode($dados,true);


foreach($dados_array as $valor){
  $produto = $valor["Produto"];
  $quantidade = $valor["Quantidade"];
  $valor = $valor["Valor unidade"];

  mysqli_commit($conn);
  
  $sql = "INSERT INTO entrada_produto (codigo, produto, quantidade, valor)
  VALUES ('$codigo','$produto','$quantidade','$valor')";
  mysqli_query($conn, $sql);

  $sql = "SELECT * FROM produto WHERE descricao = '$produto'";
  $resultado = mysqli_query($conn, $sql);

    while($row = mysqli_fetch_assoc($resultado)){
        $novoValor = $row['estoque'] + $quantidade;
        $sql = "UPDATE produto SET estoque = $novoValor WHERE descricao = '$produto'";
        $res = mysqli_query($conn, $sql);
      }

      mysqli_free_result($resultado);

}


$sql = "INSERT INTO entrada (codigo, origem, destino, emissao, recebimento)
VALUES ('$codigo','$origem','$destino','$emissao', '$recebimento')";
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
        window.location.replace('entrada.php')
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
        window.location.replace('entrada.php')
    }
    </script>
</body>
<?php
}
          
?>