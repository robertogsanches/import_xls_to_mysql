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
	 * Método que verifica registros duplicados
	  * 
	 * Para validar registros duplicados precisa buscar o campo codigoRelacaoEliminacao na tbl_relacaoeliminacao 
	 * passando os parametros numero e ano que são resctivamente Edital e Ano do xls
	 *
	 *	edital = codigoRelacaoEliminacao
	 *	unidade = codigoEntidade
	 *	classificacao = codigoClassificacaoDocumentoMeio
	 * 
	 * @return Valor Booleano TRUE para duplicado e FALSE caso não 
	 */
	private function isRegistroDuplicado($edital=null,$ano=null,$classificacao=null,$unidade=null){

		$retorno = false;

		try{
			if(!empty($edital) && !empty($ano) && !empty($classificacao) && !empty($unidade) ){
				
				$sql = "SELECT codigoRelacaoEliminacao FROM tbl_relacaoeliminacao WHERE numero = ? and ano = ?";
				$stm = $this->conexao->prepare($sql);
				$stm->bindValue(1, $edital);
				$stm->bindValue(2, $ano);
				$stm->execute();
				$dados = $stm->fetch();

				
				if(!empty($dados)){
					$codigoRelacaoEliminacao = $dados['codigoRelacaoEliminacao'];
					//Buscar na tbl_relacaoeliminacaoitem mas antes buscar os id's de classificacao e entidade

					//Precisa ir na tabela tbl_ClassificacaoDocumentoMeio buscar o campo codigoClassificacaoDocumentoMeio usando o campo classificacaoMeio como referencia
					//classificação vem no padrão 99.99.99.99 precisa remover os pontos
					$classificacao = str_replace('.','',$classificacao);
					$sql_clas = 'SELECT codigoClassificacaoDocumentoMeio FROM tbl_classificacaodocumentomeio WHERE classificacaoMeio = ? LIMIT 1';
					$stm_clas = $this->conexao->prepare($sql_clas);
					$stm_clas->bindValue(1, $classificacao);
					$stm_clas->execute();
					$dados_clas = $stm_clas->fetch();

					if(!empty($dados_clas)){
						
						$codigoClassificacaoDocumentoMeio 	= $dados_clas['codigoClassificacaoDocumentoMeio'];

						//se existir a classificacao então busca a entidade/unidade

						//Precisa ir na tabela tbl_entidade buscar o campo codigoEntidade usando o campo descricao como referencia
						$sql_uni = "SELECT codigoEntidade FROM tbl_entidade WHERE descricao LIKE ? LIMIT 1";
						$stm_uni = $this->conexao->prepare($sql_uni);
						$stm_uni->bindValue(1, $unidade);
						$stm_uni->execute();
						$dados_uni = $stm_uni->fetch();

						if(!empty($dados_uni)){
							$codigoEntidade 	= $dados_uni['codigoEntidade'];
							//se existir a unidade então busca o relacaoeliminacaoitem
								$sqlI = "SELECT codigoRelacaoEliminacaoItem FROM tbl_relacaoeliminacaoitem WHERE codigoRelacaoEliminacao = ? and codigoEntidade = ? and codigoClassificacaoDocumentoMeio = ?";
								$stmI = $this->conexao->prepare($sqlI);
								$stmI->bindValue(1, $codigoRelacaoEliminacao);
								$stmI->bindValue(2, $codigoEntidade);
								$stmI->bindValue(3, $codigoClassificacaoDocumentoMeio);
								$stmI->execute();
								$dadosI = $stmI->fetch();

								if(!empty($dadosI)){
								//registro duplicado
									$retorno = true;	
								} else {
									$retorno = false;	
								}
						} else {
							//não encontrou o nome da unidade
							$retorno = false;
						}

					} else {
						//Caso não exista a classificação LOGO não é duplicado
						$retorno = false;
					}

					


				} else {
					//NÃO ACHOU O EDITAL (relacaoeliminacao) LOGO ESSE REGISTRO NÃO É DUPLICADO
					$retorno = false;
				}

				

			}			
			
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
				if ( $chave >= 1 && !$this->isRegistroDuplicado( trim($valor[0]),trim($valor[1]),trim($valor[2]),trim($valor[5]) ) ){
					
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
			
			if ( $novo_relacaoeliminacao > 0 ){
				echo '<h6><br/>Foram criados: ', $novo_relacaoeliminacao, ' novos registros na tabela Relação Eliminação!</h6>';
			}
			
			if ( $c_erro > 0){
				echo '<h4><br/>Foram gerados: '. $c_erro. ' erros dos quais:</h4>';
			
				if ( $erro_relacaoeliminacao > 0){
					echo '<h5><br/>'. $erro_relacaoeliminacao. ' erros para inserir na tabela relacao eliminacao</h5>';
				}
			
				if ( $erro_unidade > 0){
					echo '<h5><br/>'. $erro_unidade, ' erros no campo Unidade (nome divergente)</h5>';
				}
				if ( $erro_classificacao > 0){
					echo '<h5><br/>'. $erro_classificacao. ' erros com classificação </h5>';        
				}
				
			}

			return $linhas_importadas;
		}catch(Exception $erro){
			echo 'Erro: ' . $erro->getMessage();
		}

	}
}