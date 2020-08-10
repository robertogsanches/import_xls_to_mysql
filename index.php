<?php
/* Seta configuração para não dar timeout */
ini_set('max_execution_time','-1');

/* Require com a classe de importação construída */
require_once __DIR__.'/importaplanilha.class.php';

/* Instância conexão PDO com o banco de dados */
$pdo = new PDO('mysql:host=localhost;dbname=protocolo', 'root', '');

/* Instância o objeto importação e passa como parâmetro o caminho da planilha e a conexão PDO */
$obj = new ImportaPlanilha('./teste.xlsx', $pdo);

/* Chama o método que retorna a quantidade de linhas */
echo '<h1>Quantidade de Linhas na Planilha ' , $obj->getQtdeLinhas(), '</h1><br><br>';

/* Chama o método que retorna a quantidade de colunas */
//echo 'Quantidade de Colunas na Planilha ' , $obj->getQtdeColunas(), '<br>';

/* Chama o método que inseri os dados e captura a quantidade linhas importadas */
$linhasImportadas = $obj->insertDados();

/* Imprime a quantidade de linhas importadas */
echo '<h2><br/>Foram importadas ', $linhasImportadas, ' linhas</h2>';

/* Chama o método que exibe os erros*/

/* Imprime a quantidade de linhas importadas e erros*/

//echo '<h2><br/>Foram importadas '. $linhasImportadas['insert'].' linhas</h2>';









exit;
?>
