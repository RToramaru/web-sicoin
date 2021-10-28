<?php

require "../banco_dados/conecta.php";


$sql = "SELECT setor FROM setor WHERE tipo = 'entrada' OR tipo = 'entrada/saida'";
$entradas = mysqli_query($conn, $sql);


$sql = "SELECT setor FROM setor WHERE tipo = 'saida' OR tipo = 'entrada/saida'";
$saidas = mysqli_query($conn, $sql);


$sql = "SELECT descricao FROM produto";
$produtos = mysqli_query($conn, $sql);

$itens_tabela = 0;
?>