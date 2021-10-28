<?php

require "../banco_dados/conecta.php";


$descricao = filter_input(INPUT_GET, 'descricao', FILTER_SANITIZE_STRING);

$sql = "DELETE FROM produto WHERE descricao = '$descricao'";

if (mysqli_query($conn, $sql)) {
  header('location:produto.php');
} else {
  echo "Error: " . $sql . "<br>" . mysqli_error($conn);
}

?>