<?php 
ini_set('max_execution_time','-1');
require_once "SimpleXLSX.class.php";

class ImportaPlanilha{

	// Atributo recebe a instância da conexão PDO
	private $conexao  = null;

     // Atributo recebe uma instância da classe SimpleXLSX
	private $planilha = null;

	// Atributo recebe a quantidade de linhas da planilha
	private $linhas   = null;

	// Atributo recebe a quantidade de colunas da planilha
	private $colunas  = null;

	/*
	 * Método Construtor da classe
	 * @param $path - Caminho e nome da planilha do Excel xlsx
	 * @param $conexao - Instância da conexão PDO
	 */
	public function __construct($path=null, $conexao=null){

		if(!empty($path) && file_exists($path)):
			$this->planilha = new SimpleXLSX($path);
			list($this->colunas, $this->linhas) = $this->planilha->dimension();
		else:
			echo 'Arquivo não encontrado!';
			exit();
		endif;

		if(!empty($conexao)):
			$this->conexao = $conexao;
		else:
			echo 'Conexão não informada!';
			exit();
		endif;

	}

	/*
	 * Método que retorna o valor do atributo $linhas
	 * @return Valor inteiro contendo a quantidade de linhas na planilha
	 */
	public function getQtdeLinhas(){
		//Menos uma linha que é o cabeçalho
		return $this->linhas-1;
	}

	/*
	 * Método que retorna o valor do atributo $colunas
	 * @return Valor inteiro contendo a quantidade de colunas na planilha
	 */
	public function getQtdeColunas(){
		return $this->colunas;
	}

	/*
	 * Método que verifica se o registro CPF da planilha já existe na tabela cliente
	 * @param $cpf - CPF do cliente que está sendo lido na planilha
	 * @return Valor Booleano TRUE para duplicado e FALSE caso não 
	 */
	private function isRegistroDuplicado($cpf=null){
		$retorno = false;

		try{
			if(!empty($cpf)):
				$sql = 'SELECT id FROM cliente WHERE cpf = ?';
				$stm = $this->conexao->prepare($sql);
				$stm->bindValue(1, $cpf);
				$stm->execute();
				$dados = $stm->fetchAll();

				if(!empty($dados)):
					$retorno = true;
				else:
					$retorno = false;
				endif;
			endif;

			
		}catch(Exception $erro){
			echo 'Erro: ' . $erro->getMessage();
			$retorno = false;
		}

		return $retorno;
	}

	/*
	 * Método para ler os dados da planilha e inserir no banco de dados
	 * @return Valor Inteiro contendo a quantidade de linhas importadas
	 */
	public function insertDados(){	

		$retorno_insert = array(
			'insert' => '',
			'erro' => '',
			'novo_relacaoeliminacao' => 0,
			'erro_relacaoeliminacao' => 0,
			'erro_unidade' => 0,
			'erro_classificacao' => 0
		);


		try{
			//$sql = 'INSERT INTO TABELA (edital, ano, classificacao, caixas, data_limite, unidade, data_publicacao, pagina, observacao )VALUES(?, ?, ?, ?, ?, ?, ?, ?)';

			$sql = 'INSERT INTO tbl_relacaoeliminacaoitem (codigoRelacaoEliminacao, codigoEntidade, codigoClassificacaoDocumentoMeio, codigoUsuario, dataLimite, totalCaixa, situacao, observacao) VALUES (?, ?, ?, ?, ?, ?, ?, ?)';
			$stm = $this->conexao->prepare($sql);
			

			//variavel para ver em qual linha do registro está o erro
			$linha = 1;
			//contador com as linhas que foram importadas
			$linhas_importadas = 0;
			//variavel para contar o erros de inserção
			$c_erro = 0;
			
			$novo_relacaoeliminacao = 0;
			$erro_relacaoeliminacao = 0;
			$erro_unidade = 0;
			$erro_classificacao = 0;

			foreach($this->planilha->rows() as $chave => $valor):
				if ($chave >= 1 && !$this->isRegistroDuplicado(trim($valor[2]))){
					
					$flag_erro = FALSE;

					//pega os camnpos(colunas) do xls
					$xls_edital  			= trim($valor[0]);
					$xls_ano    			= trim($valor[1]);
					$xls_classificacao     	= trim($valor[2]);
					$xls_caixas   			= trim($valor[3]);
					$xls_data_limite 		= trim($valor[4]);
					$xls_unidade 			= trim($valor[5]);
					$xls_data_publicacao 	= trim($valor[6]);
					$xls_pagina 			= trim($valor[7]);
					$xls_observacao 		= trim($valor[8]);


					//prepara os campos para o insert

						$sql0 = "SELECT codigoRelacaoEliminacao FROM tbl_relacaoeliminacao WHERE numero = ? and ano = ?";
						$stm0 = $this->conexao->prepare($sql0);
						$stm0->bindValue(1, $xls_edital);
						$stm0->bindValue(2, $xls_ano);
						$stm0->execute();
						$dados0 = $stm0->fetch();

						if(!empty($dados0)){
							$codigoRelacaoEliminacao = $dados0['codigoRelacaoEliminacao'];
						} else {
							$novo_relacaoeliminacao++;
			
							//Caso não exista o RelacaoEliminacao PAI precisa cadastrar ele
							//codigoRelacaoEliminacao id auto increment
							$sql_cad_pai = 'INSERT INTO tbl_relacaoeliminacao (
								numero, 
								ano, 
								dataPublicacao, 
								pagina
								) VALUES (?, ?, ?, ?)';

							$stm_cad_pai = $this->conexao->prepare($sql_cad_pai);
							$stm_cad_pai->bindValue(1, $xls_edital);
							$stm_cad_pai->bindValue(2, $xls_ano);
							$stm_cad_pai->bindValue(3, $xls_data_publicacao);
							$stm_cad_pai->bindValue(4, $xls_pagina);
							$retorno_cad_pai = $stm_cad_pai->execute();

							if($retorno_cad_pai == true) {
							//busca o id do pai para inserir no filho
								$codigoRelacaoEliminacao = $this->conexao->lastInsertId();
							} else {
							//erro ao cadastrar o RelacaoEliminacao Pai
								$c_erro++;
								$erro_relacaoeliminacao++;
								$flag_erro = TRUE;
								echo "<br/>Erro na linha: <b>$linha</b> Edital: $xls_edital Ano: $xls_ano";
							}							
						}

						//buscar na tab Entidade usando $xls_unidade como referencia
						$sql1 = "SELECT codigoEntidade FROM tbl_entidade WHERE descricao LIKE ? LIMIT 1";
						$stm1 = $this->conexao->prepare($sql1);
						$stm1->bindValue(1, $xls_unidade);
						$stm1->execute();
						$dados1 = $stm1->fetch();

						if(!empty($dados1)){
							$codigoEntidade 	= $dados1['codigoEntidade'];
						} else {
							//caso tenha algum nome de unidade divergente não vai salvar
							$erro_unidade++;
							$c_erro++;
							$flag_erro = TRUE;
						}

						//classificação vem no padrão 99.99.99.99 precisa remover os pontos
						$xls_classificacao = str_replace('.','',$xls_classificacao);
						$sql2 = 'SELECT codigoClassificacaoDocumentoMeio FROM tbl_classificacaodocumentomeio WHERE classificacaoMeio = ? LIMIT 1';
						$stm2 = $this->conexao->prepare($sql2);
						$stm2->bindValue(1, $xls_classificacao);
						$stm2->execute();
						$dados2 = $stm2->fetch();
		
						if(!empty($dados2)){
							$codigoClassificacaoDocumentoMeio 	= $dados2['codigoClassificacaoDocumentoMeio'];
						} else {
							$c_erro++;
							$erro_classificacao++;
							$flag_erro = TRUE;
							//Caso não exista a classificação exibir erro na tela e não inserir
							echo "Erro na linha: <b>$linha</b> - não existe a classificação: $xls_classificacao<br/>";
						}				

						$codigoUsuario 						= '101';
						$dataLimite							= $xls_data_limite;
						$totalCaixa							= $xls_caixas;
						$situacao							= 'A';
						$observacao							= $xls_observacao;

						if ($observacao == ''){
							$observacao = null;
						} 

					/*debug campos que vão ser inseridos
					echo ('f1 :'.$codigoRelacaoEliminacao.'</br>');
					echo ('f2 :'.$codigoEntidade.'</br>');
					echo ('f3 :'.$codigoClassificacaoDocumentoMeio.'</br>');
					echo ('f4 :'.$codigoUsuario.'</br>');
					echo ('f5 :'.$dataLimite.'</br>');
					echo ('f6 :'.$totalCaixa.'</br>');
					echo ('f7 :'.$situacao.'</br>');
					echo ('f8 :'.$observacao.'</br>');
					*/			
					
					if ($flag_erro == FALSE){
						//salva no banco
						$stm->bindValue(1, $codigoRelacaoEliminacao);
						$stm->bindValue(2, $codigoEntidade);
						$stm->bindValue(3, $codigoClassificacaoDocumentoMeio);
						$stm->bindValue(4, $codigoUsuario);
						$stm->bindValue(5, $dataLimite);
						$stm->bindValue(6, $totalCaixa);
						$stm->bindValue(7, $situacao);
						$stm->bindValue(8, $observacao);
						$retorno = $stm->execute();
						
						if($retorno == true) {
							$linhas_importadas++;
						}
					}	
				}
				$linha++;
			endforeach;
			//passa o total de linhas importadas e os erros
			$retorno_insert['insert'] 					= $linhas_importadas;
			$retorno_insert['erro'] 					= $c_erro;
			$retorno_insert['novo_relacaoeliminacao'] 	= $novo_relacaoeliminacao;
			$retorno_insert['erro_relacaoeliminacao'] 	= $erro_relacaoeliminacao;
			$retorno_insert['erro_unidade'] 			= $erro_unidade;
			$retorno_insert['erro_classificacao'] 		= $erro_classificacao;

			return $retorno_insert;
		}catch(Exception $erro){
			echo 'Erro: ' . $erro->getMessage();
		}

	}
}