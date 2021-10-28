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
    <title>Nota de Saída</title>
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

    <h1 class="container-fluid interno">Nota de Saída</h1>

    <form action="saida_insert.php" method="post">

        <div class="container-fluid">
            <div class="interno">

                <div class="form-row">
                    <div class="col-sm-6 my-1">
                        <label>código</label>
                        <input type="text" id="codigo" name="codigo" class="form-control" required>
                    </div>

                    <div class="col-sm-6 my-1">
                        <label>Natureza da operação</label>
                        <select name="natureza" id="natureza" class="form-control">
                            <option value="consumo">Remessa para consumo</option>
                            <option value="venda">Venda</option>
                            <option value="doacao">Doação</option>
                            <option value="transferencia">Transferência</option>
                            <option value="estorno">Estorno</option>
                        </select>
                    </div>

                    <div class="col-sm-6 my-1">
                        <label>Setor de destino</label>
                        <select name="destino" id="destino" class="form-control">
                            <?php
                            while($linhas_saidas = mysqli_fetch_assoc($saidas)){
                        ?>
                            <option value="<?php echo $linhas_saidas['setor'];?>">
                                <?php echo $linhas_saidas['setor'];?>
                            </option>
                            <?php }?>
                        </select>
                    </div>

                    <div class="col-sm-6 my-1">
                        <label>CNPJ:</label>
                        <input type="text" id="cnpj" name="cnpj" class="form-control">
                    </div>


                    <div class="col-sm-6 my-1">
                        <label>Data de emissão:</label>
                        <input type="date" id="emissao" name="emissao" class="form-control" required>
                    </div>

                </div>

                <div class="form-row">
                    <div class="col-sm-3 my-1">
                        <label>Produto:</label>

                        <select name="produto" id="produto" class="form-control">
                            <?php
                            while($linhas_produtos = mysqli_fetch_assoc($produtos)){
                        ?>
                            <option value="<?php echo $linhas_produtos['descricao'];?>">
                                <?php echo $linhas_produtos['descricao'];?>

                            </option>
                            <?php }?>
                        </select>
                    </div>
                    <div class="col-sm-3 my-1">
                        <label>Quantidade:</label>
                        <input type="number" id="quantidade" name="quantidade" class="form-control" required>
                    </div>

                    <div class="col-sm-3 my-1">
                        <label>Preço unidade:</label>
                        <input type="number" id="unidade" name="unidade" class="form-control" required>
                    </div>


                    <input type="hidden" id="dados" name="dados" value="">

                    <div class="col-sm-3 my-1">
                        <a id="adicionar" class="btn btn-sm btn-success" onclick="adicionaLinha('tbl','dados')">
                            Adicionar</a>
                    </div>
                </div>


                <table id="tbl" class="table table-hover" name="tbl">
                    <tr>
                        <td>Produto</td>
                        <td>Quantidade</td>
                        <td>Valor unidade</td>
                        <td>Valor Total</td>
                        <td>Excluir</td>
                    </tr>
                </table>

                <button type="submit" class="btn btn-success btn-lg float-right">Cadastrar</button>
            </div>
        </div>
    </form>

    <h1 class="container-fluid interno">Total: <span id="total">0</span></h1>

    <div class="modal fade" id="myModal" role="dialog">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <h4 class="modal-title">Erro</h4>
                </div>
                <div class="modal-body">
                    Impossível valor negativo.
                </div>
                <div class="modal-footer">
                    <input class="btn btn-primary" data-dismiss="modal" value="Fechar">
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


    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/table-to-json@1.0.0/lib/jquery.tabletojson.min.js"
        integrity="sha256-H8xrCe0tZFi/C2CgxkmiGksqVaxhW0PFcUKZJZo1yNU=" crossorigin="anonymous"></script>

    <script>
    function openModal() {
        jQuery.noConflict();
        $('#myModal').modal('show')
    }

    function adicionaLinha(idTabela, idDados) {

        var produto = document.querySelector("select[name='produto']").value
        var quantidade = document.querySelector("input[name='quantidade']").value
        var unidade = document.querySelector("input[name='unidade']").value
        var total = document.getElementById("total").innerHTML

        if (quantidade < 0 || unidade < 0) {
            openModal()
        } else {
            if (quantidade != "" && unidade != "") {

                var tabela = document.getElementById(idTabela)
                var numeroLinhas = tabela.rows.length
                var linha = tabela.insertRow(numeroLinhas)
                var celula1 = linha.insertCell(0)
                var celula2 = linha.insertCell(1)
                var celula3 = linha.insertCell(2)
                var celula4 = linha.insertCell(3)
                var celula5 = linha.insertCell(4)
                celula1.innerHTML = produto
                celula2.innerHTML = quantidade
                celula3.innerHTML = unidade
                celula4.innerHTML = unidade * quantidade
                celula5.innerHTML =
                    "<button onclick='removeLinha(this)' class='btn btn-sm'><i class='material-icons'>delete</i></button>"

                var dado = document.getElementById(idTabela, idDados)
                var tableJson = $('#tbl').tableToJSON()
                var jsonString = JSON.stringify(tableJson)
                dados.value = jsonString
                total = parseFloat(total) + (quantidade * unidade)
                document.getElementById("total").innerHTML = total

            }
        }

    }

    function removeLinha(linha) {
        var i = linha.parentNode.parentNode.rowIndex
        var total = document.getElementById("total").innerHTML
        var mytable  =  document.getElementById('tbl')
        var objCells = mytable.rows.item(i).cells
        var valor = objCells.item(3).innerHTML
        total = parseFloat(total) - parseFloat(valor)
        document.getElementById("total").innerHTML = total
            
        document.getElementById('tbl').deleteRow(i)
        var tableJson = $('#tbl').tableToJSON()
        var jsonString = JSON.stringify(tableJson)
        dados.value = jsonString
    }
    </script>


</body>

</html>