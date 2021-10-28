<?php

require "../banco_dados/conecta.php";

$setor = $_POST['setor'];
$responsavel = $_POST['responsavel'];
$tipo = $_POST['tipo'];
$cnpj = $_POST['cnpj'];


$sql = "UPDATE setor SET responsavel = '$responsavel', tipo = '$tipo', cnpj = '$cnpj' WHERE setor = '$setor'";

if (mysqli_query($conn, $sql)) {
  header('location:setor.php');
} else {
  echo "Error: " . $sql . "<br>" . mysqli_error($conn);
}

?>