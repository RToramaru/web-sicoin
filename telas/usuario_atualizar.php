<?php

require "../banco_dados/conecta.php";

$nome = $_POST['nome'];
$endereco = $_POST['endereco'];
$telefone = $_POST['telefone'];
$usuario = $_POST['usuario'];
$tipo = $_POST['tipo'];


$sql = "UPDATE usuario SET endereco = '$endereco', telefone = '$telefone', usuario = '$usuario',
tipo = '$tipo' WHERE nome = '$nome'";

if (mysqli_query($conn, $sql)) {
  header('location:usuario.php');
} else {
  echo "Error: " . $sql . "<br>" . mysqli_error($conn);
}

?>