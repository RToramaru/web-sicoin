<?php
session_start();
if(!$_SESSION['tipo'] || $_SESSION['tipo'] == 'padrao'){
    header('location:login.php');
} else{
    require "setor_select.php";
} 
?>

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <title>Setor</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
    <link rel="stylesheet" href="../css/estilo.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet" />

</head>

<body>
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

    <h1 class="container-fluid interno">Setor</h1>

    <form action="setor_insert.php" method="post">
        <div class="container-fluid">
            <div class="interno">

                <div class="form-row">
                    <div class="col-sm-6 my-1">
                        <label>Setor:</label>
                        <input type="text" id="setor" name="setor" class="form-control" required>
                    </div>

                    <div class="col-sm-6 my-1">
                        <label>Responsável:</label>
                        <input type="text" id="responsavel" name="responsavel" class="form-control" required>
                    </div>



                    <div class="col-sm-6 my-1">
                        <label>Tipo:</label>
                        <select name="tipo" id="tipo" class="form-control">
                            <option value="entrada">Entrada</option>
                            <option value="saida">Saída</option>
                            <option value="entrada/saida">Entrada/Saída</option>
                        </select>
                    </div>

                    <div class="col-sm-6 my-1">
                        <label>CNPJ:</label>
                        <input type="text" id="cnpj" name="cnpj" class="form-control" pattern="[0-9]+([.][0-9]+)?"
                            min="0" step="any">
                    </div>
                </div>
                <button type="submit" class="btn btn-success btn-lg float-right">Cadastrar</button>
            </div>
        </div>
        </div>
    </form>


    <div class="container-fluid">

        <div class="tabela">
            <?php
        if($rows > 0){?>
            <table class="table table-hover" id="minha_tabela">

                <thead>
                    <tr>
                        <th>Setor</th>
                        <th>Responsável</th>
                        <th>Tipo</th>
                        <th>CNPJ</th>
                        <th>Editar</th>
                        <th>Excluir</th>
                    </tr>
                </thead>
                <tbody>
                    <?php
            while($row = mysqli_fetch_assoc($resultado)){

                ?>

                    <tr>
                        <td><?php echo $row['setor'];?></td>
                        <td><?php echo $row['responsavel'];?></td>
                        <td><?php echo $row['tipo'];?></td>
                        <td><?php echo $row['cnpj'];?></td>

                        <td>
                            <a class="btn btn-sm" data-toggle="modal" data-target="#modalSetor"
                                data-setor="<?php echo $row['setor'];?>"
                                data-responsavel="<?php echo $row['responsavel'];?>"
                                data-tipo="<?php echo $row['tipo'];?>" data-cnpj="<?php echo $row['cnpj'];?>">
                                <i class="material-icons">create</i>
                            </a>
                        </td>


                        <td><a href="setor_delete.php?setor=<?php echo $row['setor'];?>" class="btn btn-sm"
                                data-confirmar="Tem certeza que deseja excluir o item selecionado ?">
                                <i class="material-icons">delete</i></button></a></td>
                    </tr>
                    <?php }
        ?>
                </tbody>
            </table>
            <?php }
    ?>

        </div>
    </div>



    <div id="modalSetor" class="modal fade bd-example-modal-lg" role="dialog">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h4 class="modal-title">Editar setor</h4>
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                </div>
                <div class="modal-body">
                    <form action="setor_atualizar.php" method="post">
                        <div class="container-fluid">
                            <div class="form-row">
                                <div class="col-sm my-1">
                                    <label>Setor:</label>
                                    <input type="text" id="setor" name="setor" class="form-control" readonly>
                                </div>

                                <div class="col-sm my-1">
                                    <label>Responsável:</label>
                                    <input type="text" id="responsavel" name="responsavel" class="form-control">
                                </div>
                            </div>

                            <div class="form-row">

                                <div class="col-sm my-1">
                                    <label>Tipo:</label>
                                    <select name="tipo" id="tipo" class="form-control">
                                        <option value="entrada">Entrada</option>
                                        <option value="saida">Saída</option>
                                        <option value="entrda/saida/Saida">Entrada/Saída</option>
                                    </select>
                                </div>
                            </div>

                            <div class="form-row">
                                <div class="col-sm my-1">
                                    <label>CNPJ:</label>
                                    <input type="text" id="cnpj" name="cnpj" class="form-control">
                                </div>
                            </div>
                            <button id="editar" type="submit" class="btn btn-success float-right">Atualizar</button>
                        </div>

                    </form>
                </div>
            </div>
        </div>
    </div>


    <div id="rodape">
        <footer>
            <div class="container-fluid interno">
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
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>
    <script src="../js/script.js"></script>

    <link rel="stylesheet" href="https://cdn.datatables.net/1.10.10/css/jquery.dataTables.min.css">
    <script src="https://cdn.datatables.net/1.10.10/js/jquery.dataTables.min.js"></script>


    <script>
    $(document).ready(function() {

        $('#minha_tabela').DataTable({
            "language": {
                "lengthMenu": "Mostrando _MENU_ registros por página",
                "zeroRecords": "Nada encontrado",
                "info": "Mostrando página _PAGE_ de _PAGES_",
                "infoEmpty": "Nenhum registro disponível",
                "infoFiltered": "(filtrado de _MAX_ registros no total)",
                "search": "Buscar:"
            }
        })
    })
    </script>

</body>

</html>