<?php
session_start();
if(!$_SESSION['tipo']){
    header('location:login.php');
}else {
    require "option_select.php";
}  
?>

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <title>Relatório</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
    <link rel="stylesheet" href="../css/estilo.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet" />

</head>

<body>
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

    <h1 class="container-fluid interno centro">Relatório</h1>

    <form action="gerar_pdf.php" method="post">
        <div class="posicao">
            <div class="container col-md-5">
                <div class="interno">
                    <div class="col-md">
                        <label>Tipo de nota:</label>
                        <select name="tipo" id="tipo" class="form-control">
                            <option value="entrada">Entrada</option>
                            <option value="saida">Saída</option>
                        </select>
                    </div>

                    <div class="col-md">
                        <label>Código da nota:</label>
                        <input type="text" id="codigo" name="codigo" class="form-control">
                    </div>

                    <div class="col-md">
                        <label>Produto:</label>

                        <select name="produto" id="produto" class="form-control">
                            <option value=""></option>
                            <?php
                            while($linhas_produtos = mysqli_fetch_assoc($produtos)){
                        ?>
                            <option value="<?php echo $linhas_produtos['descricao'];?>">
                                <?php echo $linhas_produtos['descricao'];?>

                            </option>
                            <?php }?>
                        </select>
                    </div>

                    <div class="col-md">
                        <label>Setor Destino</label>
                        <select name="setor" id="setor" class="form-control">
                            <option value=""></option>
                            <?php
                            while($linhas_saidas = mysqli_fetch_assoc($saidas)){
                        ?>
                            <option value="<?php echo $linhas_saidas['setor'];?>">
                                <?php echo $linhas_saidas['setor'];?>
                            </option>
                            <?php }?>
                        </select>
                    </div>

                    <div class="col-md">
                        <label>Data de:</label>
                        <input type="date" id="inicio" name="inicio" class="form-control">
                    </div>

                    <div class="col-md">
                        <label>Data até:</label>
                        <input type="date" id="fim" name="fim" class="form-control">
                    </div>
                    <div class="col-md">
                        <button type="submit" class="btn-success btn-lg col-md botao">Gerar Relatório</button>
                    </div>
                </div>
            </div>
        </div>
    </form>

    <div id="rodape">
        <footer>
            <div class="container-fluid interno centro">
                <div class="row">
                    <div class="col-sm-12">
                        <div class="copyright-box">
                            <p class="copyright">&copy; Desenvolvido por <strong>Rafael Almeida Soares </strong> para o IFNMG</p>
                        </div>
                    </div>
                </div>
            </div>
        </footer>
    </div>

    <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>

</body>

</html>