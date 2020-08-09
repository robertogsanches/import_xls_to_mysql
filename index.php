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

/* Chama o método que exibe os erros*/

/* Imprime a quantidade de linhas importadas e erros*/
echo '<h2><br/>Foram importadas '. $linhasImportadas['insert'].' linhas</h2>';

if ( $linhasImportadas['novo_relacaoeliminacao'] > 0 ){
    echo '<h3><br/>Foram criados: ', $linhasImportadas['novo_relacaoeliminacao'], ' novos registros na tabela Relação Eliminação!</h3>';
}

if ( $linhasImportadas['erro'] > 0){
    echo '<h4><br/>Foram gerados: '. $linhasImportadas['erro']. ' erros dos quais:</h4>';

    if ( $linhasImportadas['erro_relacaoeliminacao'] > 0){
        echo '<h5><br/>'. $linhasImportadas['erro_relacaoeliminacao']. ' erros para inserir na tabela relacao eliminacao</h5>';
    }

    if ( $linhasImportadas['erro_unidade'] > 0){
        echo '<h5><br/>'. $linhasImportadas['erro_unidade'], ' erros no campo Unidade (nome divergente)</h5>';
    }
    if ( $linhasImportadas['erro_classificacao'] > 0){
        echo '<h5><br/>'. $linhasImportadas['erro_classificacao']. ' erros com classificação </h5>';        
    }
    
}







exit;
?>
