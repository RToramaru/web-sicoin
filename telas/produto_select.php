<?php

require "../banco_dados/conecta.php";


$sql = "SELECT * FROM produto";
$resultado = mysqli_query($conn, $sql);
$rows = mysqli_num_rows($resultado);
?>