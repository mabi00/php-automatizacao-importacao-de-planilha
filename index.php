<?php 
session_start();


// Turn off all error reporting
error_reporting(0);

require_once "PHPExcel/Classes/PHPExcel.php";

    $tmpfname = "teste3.xlsx";
    $excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
    $excelObj = $excelReader->load($tmpfname);
    $worksheet = $excelObj->getSheet(1);
    $lastRow = $worksheet->getHighestRow();
?>

    <!DOCTYPE html>
    <html lang="pt-br">

    <head>
        <!-- Required meta tags -->
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <!-- Bootstrap CSS -->
        <link rel="stylesheet" href="bower_componets/bootstrap/css/bootstrap.min.css">

        <link rel="stylesheet" type="text/css" href="bower_componets/bootstrap/estilo.css">
        <link href="bower_componets/bootstrap/DataTables/DataTables-1.10.18/css/jquery.dataTables.css" rel="stylesheet">

        <title>Processar Planilha NFE</title>
    </head>

    <body>
        <h1>Planilha NFE</h1>
        <div class="container">
            <br>
            <?php
            if(isset($_SESSION['msg'])){
                echo $_SESSION['msg'];
                unset($_SESSION['msg']);
            }
            ?>
                <header class="row">

                </header>

                <div>

                    <div class="col-md-12">
                        <form action="#" class="form-vertical">
                            <div class="form-group">
                                <div class="input-group">
                                    <div class="custom-file">
                                        <input type="file" id="arquivo" name="arquivo" class="custom-file-input">
                                        <label id="mostraArquivo" class="custom-file-label" data-browse="Buscar"> Selecionar arquivo </label>
                                    </div>
                                    <div class="input-group-append">

                                        <button type="submit" id="processar" class="btn btn-outline-success">Processar</button>
                                    </div>

                                </div>

                            </div>
                            <div class="form-group">
                                <div class="progress progress-striped active">
                                    <div class="progress-bar" style="width: 0%">
                                    </div>
                                </div>
                            </div>
                        </form>
                    </div>

                </div>
                <!-- Footer -->

        </div>
        <footer>

        </footer>
        <!-- Footer -->
        <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js" integrity="sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1" crossorigin="anonymous"></script>
        <script src="bower_componets/jquery/bootstrap.min.js"></script>
        <script src="bower_componets/bootstrap/DataTables/DataTables-1.10.18/js/jquery.dataTables.min.js" type="text/javascript"></script>

        <script type="text/javascript">
            $(document).on('submit', 'form', function(e) {
                e.preventDefault();
                //Receber os dados
                $form = $(this);
                var formdata = new FormData($form[0]);

                //Criar a conexao com o servidor
                var request = new XMLHttpRequest();

                //Progresso do Upload
                request.upload.addEventListener('progress', function(e) {
                    var percent = Math.round(e.loaded / e.total * 100);
                    $form.find('.progress-bar').width(percent + '%').html(percent + '%');
                    $('#processar').attr('disabled', 'disabled');

                });

                //Upload completo limpar a barra de progresso
                request.addEventListener('load', function(e) {
                    $form.find('.progress-bar').addClass('bg-success').html('Upload de Planilha Completo...');
                    $('#processar').removeAttr('disabled');

                    //Atualizar a página após o upload completo
                    setTimeout("window.open(self.location, '_self');", 1000);
                });

                //Arquivo responsável em fazer o upload da imagem
                request.open('post', 'processaPlanilha.php');
                request.send(formdata);
            });
            $(function() {
                $('#arquivo').change(function() {
                    var arquivo = $(this).val().replace("C:\\fakepath\\", "");
                    $('#mostraArquivo').html(arquivo);
                });
            });
        </script>
    </body>

    </html>