 <?php

    define('DB_HOST' , "death-star");
    define('DB_USER', "mabiobanco");
    define('DB_PASSWORD', "mabio123");
    define('DB_NAME', "DIGITAL");
    //define('DB_DRIVER', "odbc:Driver={SQL Server}");
    define('DB_DRIVER', "sqlsrv");

require_once("Conexao.php");

function open_database()
{
    try {
        $conn = Conexao::getConnection();
        return $conn;
    }
    catch (Exception $e) {
        echo $e->getMessage();
        return null;
    }
}

function close_database($conn)
{
    try {
        $conn = null;
        
    }
    catch (Exception $e) {
        echo $e->getMessage();
    }
}



              
