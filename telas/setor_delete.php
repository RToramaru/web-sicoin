<?php

require "../banco_dados/conecta.php";


$setor = filter_input(INPUT_GET, 'setor', FILTER_SANITIZE_STRING);

$sql = "DELETE FROM setor WHERE setor = '$setor'";

if (mysqli_query($conn, $sql)) {
  header('location:setor.php');
} else {
  echo "Error: " . $sql . "<br>" . mysqli_error($conn);
}

?>