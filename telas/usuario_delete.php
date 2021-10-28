<?php

require "../banco_dados/conecta.php";


$nome = filter_input(INPUT_GET, 'nome', FILTER_SANITIZE_STRING);

$sql = "DELETE FROM usuario WHERE nome = '$nome'";

if (mysqli_query($conn, $sql)) {
  header('location:usuario.php');
} else {
  echo "Error: " . $sql . "<br>" . mysqli_error($conn);
}

?>