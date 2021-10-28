<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <title>Acesso</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
    <link rel="stylesheet" href="../css/estilo.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet" />

</head>

<body>


    <h1 class="container interno centro espaco">Acesso</h1>

    <form action="login_select.php" method="post">

        <div class="posicao">
            <div class="container col-md-5">
                <div class="interno">
                    <div class="col-md">
                        <label>Usuário</label>
                        <input type="text" id="usuario" name="usuario" class="form-control" required>
                    </div>

                    <div class="col-md">
                        <label>Senha:</label>
                        <input type="password" id="senha" name="senha" class="form-control" required>
                    </div>

                    <div class="col-md">
                        <button type="submit" class="btn-success btn-lg col-md botao">Acessar</button>
                    </div>
                </div>
            </div>
        </div>
    </form>

</body>

</html>