<?php
require "../banco_dados/conecta.php";
session_start();
$usuario = $_POST['usuario'];
$senha = $_POST['senha'];

$sql = "SELECT * FROM usuario WHERE usuario = '$usuario' and senha = '$senha'";
$resultado = mysqli_query($conn, $sql);
$rows = mysqli_num_rows($resultado);

if($rows == 1){
    while($row = mysqli_fetch_assoc($resultado)){
        
        if($row['tipo'] == 'padrao'){
            $_SESSION['tipo'] = 'padrao';
        }else{
            $_SESSION['tipo'] = 'administrador';
        }
    }
    header('location:index.php');
} else{
    header('location:login.php');
}

?>