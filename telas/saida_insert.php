<?php

require "../banco_dados/conecta.php";



$codigo = $_POST['codigo'];
$natureza = $_POST['natureza'];
$destino = $_POST['destino'];
$cnpj = $_POST['cnpj'];
$emissao = $_POST['emissao'];
$dados = $_POST['dados'];
$valida = true;

$dados_array = json_decode($dados,true);


foreach($dados_array as $valor){
  $produto = $valor["Produto"];
  $quantidade = $valor["Quantidade"];
  $valor = $valor["Valor unidade"];

  mysqli_commit($conn);
  
  $sql = "INSERT INTO saida_produto (codigo, produto, quantidade, valor)
  VALUES ('$codigo','$produto','$quantidade','$valor')";
  mysqli_query($conn, $sql);
  
  $sql = "SELECT * FROM produto WHERE descricao = '$produto'";
  $resultado = mysqli_query($conn, $sql);

  while($row = mysqli_fetch_assoc($resultado)){
    if(($row['estoque'] - $quantidade) < 0){
      $valida = false;
    }else{
      $novoValor = $row['estoque'] - $quantidade;
      $sql = "UPDATE produto SET estoque = $novoValor WHERE descricao = '$produto'";
      $res = mysqli_query($conn, $sql);
    } 
  }
  if($valida){
    $sql = "INSERT INTO saida_produto (codigo, produto, quantidade, valor)
    VALUES ('$codigo','$produto','$quantidade','$valor')";
    mysqli_query($conn, $sql);
  } 
}

if($valida){
  $sql = "INSERT INTO saida (codigo, natureza, destino, cnpj, emissao)
  VALUES ('$codigo','$natureza','$destino','$cnpj', '$emissao')";
  mysqli_autocommit($conn,true);
} else{
  mysqli_rollback($conn);
  $valida = false;
}

if ($valida) {
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
        window.location.replace('saida.php')
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
        window.location.replace('saida.php')
    }
    </script>
</body>
<?php
  }
      
  ?>