<?php
session_start();
if(!$_SESSION['tipo']){
    header('location:login.php');
}
?>

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <title>Principal</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
    <link rel="stylesheet" href="../css/estilo.css">     

</head>

<body class="fundo">
    <?php if($_SESSION['tipo'] == 'administrador') {?>
    <nav id="navbar" class="navbar navbar-expand-lg">
        <img src="../imagens/SICOIN.svg" alt="triangle with all three sides equal" height="100%" width="100px" />
        <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavAltMarkup"
            aria-controls="navbarNavAltMarkup" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNavAltMarkup">
            <div class="nav navbar-nav ml-auto">
                <a id="item" class="nav-item nav-link active" href="usuario.php">Usuário
                    <a id="item" class="nav-item nav-link" href="setor.php">Setor</a>
                    <a id="item" class="nav-item nav-link" href="produto.php">Produto</a>
                    <a id="item" class="nav-item nav-link" href="entrada.php">Nota de entrada</a>
                    <a id="item" class="nav-item nav-link" href="saida.php">Nota de saída</a>
                    <a id="item" class="nav-item nav-link" href="relatorio.php">Relatórios</a>
                    <a id="item" class="nav-item nav-link" href="sair.php">Sair</a>
            </div>
        </div>
    </nav>

    <?php } else {?>

    <nav id="navbar" class="navbar navbar-expand-lg">
        <img src="../imagens/SICOIN.svg" alt="triangle with all three sides equal" height="100%" width="100px" />
        <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavAltMarkup"
            aria-controls="navbarNavAltMarkup" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNavAltMarkup">
            <div class="nav navbar-nav ml-auto">
                <a id="item" class="nav-item nav-link" href="entrada.php">Nota de entrada</a>
                <a id="item" class="nav-item nav-link" href="saida.php">Nota de saída</a>
                <a id="item" class="nav-item nav-link" href="relatorio.php">Relatórios</a>
                <a id="item" class="nav-item nav-link" href="sair.php">Sair</a>
            </div>
        </div>
    </nav>


    <?php } ?>

    <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>

</body>

</html>