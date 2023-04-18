<?php

class Conexao{

	private static $connection;

	

	public static function getConnection(){

		$pdoConfig = DB_DRIVER . ":" . "Server=" . DB_HOST . ";";
		$pdoConfig .= "Database=".DB_NAME.";";

		try{
			if (!isset($connection)) {
				$connection = new PDO($pdoConfig, DB_USER, DB_PASSWORD);
				//$connection = new PDO("sqlsrv:Server=servbd;Database=CORPORE", "teste", "1234");
				$connection->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
			}
			return $connection;

		} catch(PDOException $e){
			$mensagem = "Drivers disponiveis: " . implode(",", PDO::getAvailableDrivers());
			$mensagem .= "\nErro: " . $e->getMessage();
			throw new Exception($mensagem);
			
		}
	}
}

