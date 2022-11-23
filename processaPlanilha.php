<?php 
require_once "PHPExcel/Classes/PHPExcel.php";
require_once "database.php";

session_start();

if (isset($_FILES)) {
    $tmpfname = $_FILES['arquivo']['tmp_name'];
    //$tmpfname = "teste3.xlsx";
    $excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
    $excelObj = $excelReader->load($tmpfname);
    $worksheet = $excelObj->getSheet(1);
    $lastRow = $worksheet->getHighestRow();
}else{
       $_SESSION['msg'] = '<div class="alert alert-danger alert-dismissible fade show" role="alert">
                   <strong>Erro!</strong>
                  <button type="button" class="close" data-dismiss="alert" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                  </button>
                </div>';
}

    function formataData($data){
      $dateTime = new DateTime("1899-12-30 + $data days");

      return $dateTime->format("Y/m/d");
    }

     $database = open_database();

for ($row=2; $row <= $lastRow ; $row++) { 

      $dataEmissao[$row] = formataData( (int) $worksheet->getCell('W'.$row)->getValue());
      $competenciaManifestacao[$row] = formataData( (int) $worksheet->getCell('AG'.$row)->getValue());
      $dataArquivei[$row] = formataData( (int) $worksheet->getCell('AI'.$row)->getValue());
      $dataAutorizacao[$row] = formataData( (int) $worksheet->getCell('BF'.$row)->getValue());
      $itemDescricao = str_replace("'", " ",$worksheet->getCell('CB'.$row)->getValue() );
      $cidadeOrigem = str_replace("'", " ",$worksheet->getCell('AM'.$row)->getValue() );
        $dadosAdicionaisAspas = str_replace("'", " ",$worksheet->getCell('BQ'.$row)->getValue() );
            $dadosAdicionais = str_replace(";", " ",$dadosAdicionaisAspas);

        $dadosAdicionaisFiscoAspas = str_replace("'", " ",$worksheet->getCell('BS'.$row)->getValue() );
            $dadosAdicionais = str_replace(";", " ",$dadosAdicionaisFiscoAspas);

      $itemUnidadeAspas = str_replace("'", " ",$worksheet->getCell('CG'.$row)->getValue() ); 
        if (strlen($itemUnidadeAspas) > 5){
            $itemUnidade = substr($itemUnidadeAspas, 0, 5);
        }

      $sql = "INSERT INTO [dbo].[ZFISNOTAS]
           ([SITUACAO]
           ,[DESTINATARIO]
           ,[CNPJDESTINATARIO]
           ,[CPFDESTINATARIO]
           ,[IEDESTINATARIO]    
           ,[TIPOIE]

           ,[CEPDESTINATARIO]
           ,[RUADESTINATARIO]
           ,[NUMDESTINATARIO]
           ,[BAIRRODESTINATARIO]
           ,[EMITENTE]
           ,[CNPJEMITENTE]
           ,[EMITENTEPF]
           ,[CPFEMITENTE]
           ,[IEEMITENTE]
           ,[IESTREMETENTE]
           ,[CEPEMITENTE]
           ,[RUAEMITENTE]
           ,[NUMEMITENTE]
           ,[BAIRROEMITENTE]
           ,[CCE]
           ,[CHAVEACESSO]
           ,[DATAEMISSAO]
           ,[NUMERO]
           ,[MODELO]
           ,[SERIE]
           ,[UFORIGEM]
           ,[UFDESTINO]
           ,[TIPO]
           ,[VALORTOTALDOCUMENTO]
           ,[STATUS]
           ,[MANIFESTACAO]
           ,[COMPETENCIAMANIFESTACAO]
           ,[NSU]
           ,[DATAARQUIVEI]
           ,[ORIGEM]
           ,[NATUREZAOPERACAO]
           ,[FINALIDADENFE]
           ,[CIDADEORIGEM]
           ,[CIDADEDESTINO]
           ,[VALORTOTALPRODUTOS]
           ,[ICMSST]
           ,[IPI]
           ,[ICMS]
           ,[ICMSDESONERADO]
           ,[ICMSFCPUFDESTINO]
           ,[ICMSUFDESTINO]
           ,[ICMSUFREMETENTE]
           ,[COFINS]
           ,[II]
           ,[PIS]
           ,[VALORTOTALFRETE]
           ,[VALORTOTALSEGURO]
           ,[VALORTOTALDESCONTO]
           ,[OUTRASDESPESAS]
           ,[PROTOCOLADA]
           ,[PROTOCOLO]
           ,[DATAAUTORIZACAO]
           ,[NUMEROFATURA]
           ,[VALORFATURA]
           ,[DESCONTOFATURA]
           ,[VALORLIQUIDOFATURA]
           ,[TRANSPORTADOR]
           ,[CNPJTRANSPORTADOR]
           ,[MOTORISTA]
           ,[PLACA]
           ,[PESOLIQUIDO]
           ,[PESOBRUTO]
           ,[DADOSADICIONAIS]
           ,[ETIQUETAS]   
           ,[DADOSADICIONAISFISCO]

           ,[VALORTOTALFCP]
           ,[VALORTOTALFCPST]
           ,[VALORTOTALFCPRETIDO]
           ,[MEIOPGTO]
           ,[BASEICMS]
           ,[BASEICMSST]
           ,[FRETE]
           ,[ICODIGO]
           ,[IDESCRICAO]
           ,[IEAN]
           ,[INCM]
           ,[ICEST]
           ,[ICFOP]
           ,[IUNIDADE]
           ,[IQUANTIDADE]
           ,[IVALORUNITARIO]
           ,[IVALORTOTALBRUTO]
           ,[INUMPEDIDOCOMPRA]
           ,[IITEMPEDIDOCOMPRA]
           ,[IDESCONTO]
           ,[IVALORFRETE]
           ,[ISEGURO]
           ,[IOUTRASDESPESAS]
           ,[IICMS]
           ,[IALIQICMS]
           ,[IBCICMSUFDESTINO]
           ,[IPERCENTUALICMSFCPUFDESTINO]
           ,[IALIQINTERNAUFDESTINO]
           ,[IALIQICMSINTERESTADUAL]
           ,[IPERCENTUALPARTILHAICMSINTERESTADUAL]
           ,[IICMSFCOUFDESTINO]
           ,[IICMSINTERESTADUALUFDESTINO]
           ,[IICMSINTERESTADUALUFORIGEM]
           ,[IBCICMSST]
           ,[IICMSST]
           ,[IALIQICMSST]
           ,[ICODIGOSITUACAOOPERACAO]
           ,[IALICREDITOICMSSN]
           ,[IVALORCREDITOICMSSN]
           ,[IBCISSQN]
           ,[IISS]
           ,[IALIQISS]
           ,[IALIQIPI]
           ,[IIPI]
           ,[ICSTIPI]
           ,[IBCPIS]
           ,[IALIQPIS]
           ,[IPIS]
           ,[ICSTPIS]
           ,[IBCCOFINS]
           ,[IALIQCOFINS]
           ,[ICOFINS]
           ,[ICSTCOFINS]
           ,[IBCFCP]
           ,[IPERCENTUALFCP]
           ,[IFCP]
           ,[INUMEROITEM]
           ,[IINFOADICIONAISPRODUTO]
           ,[IBCFCPST]
           ,[IPERCENTUALFCPST]
           ,[IFCPST]
           ,[ISITUACAOTRIBUTARIA]
           ,[IORIGEMMERCADORIA]
           ,[IBCIPI]
           ,[IBCICMS]
           ,[IBCICMSSTRETIDO]
           ,[IVALORICMSSTRETIDO]
           ,[IICMSDESONERADO]
           ,[IMOTIVODESONERACAO]
           ,[IBCFCPSTRETIDO]
           ,[IPERCENTUALFCPSTRETIDO]
           ,[IFCPSTRETIDO]
           ,[IPERCENTUALADICIONADOICMSST]
           ,[IGRUPOICMSSN])
            VALUES (";
           $sql .= "'". $worksheet->getCell('A'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('B'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('C'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('D'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('E'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('F'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('G'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('H'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('I'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('J'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('K'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('L'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('M'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('N'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('O'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('P'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('Q'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('R'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('S'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('T'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('U'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('V'.$row)->getValue()."'";
           $sql .= ", '".$dataEmissao[$row]."'";
           $sql .= ",'". $worksheet->getCell('X'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('Y'.$row)->getValue()."'";
           $sql .= ",'". $worksheet->getCell('Z'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('AA'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('AB'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('AC'.$row)->getValue()."'";
           $sql .= ",". $worksheet->getCell('AD'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('AE'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('AF'.$row)->getValue()."'";
           $sql .= ",".$competenciaManifestacao[$row];
           $sql .= ",'".$worksheet->getCell('AH'.$row)->getValue()."'";
           $sql .= ",'".$dataArquivei[$row]."'";
           $sql .= ",'".$worksheet->getCell('AJ'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('AK'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('AL'.$row)->getValue()."'";
           $sql .= ",'".$cidadeOrigem."'";
           $sql .= ",'".$worksheet->getCell('AN'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('AO'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AP'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AQ'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AR'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AS'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AT'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AU'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AV'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AW'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AX'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AY'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('AZ'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BA'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BB'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BC'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('BD'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('BE'.$row)->getValue()."'";
           $sql .= ",'".$dataAutorizacao[$row]."'";
           $sql .= ",'".$worksheet->getCell('BG'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('BH'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BI'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BJ'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('BK'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('BL'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('BM'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('BN'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('BO'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('BP'.$row)->getValue()."'";
           $sql .= ",'".$dadosAdicionais."'";
           $sql .= ",'".$worksheet->getCell('BR'.$row)->getValue()."'";
           $sql .= ",'".$dadosAdicionaisFisco."'";
           $sql .= ",".$worksheet->getCell('BT'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BU'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BV'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('BW'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('BX'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('BY'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('BZ'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CA'.$row)->getValue()."'";
           $sql .= ",'".$itemDescricao."'";
           $sql .= ",'".$worksheet->getCell('CC'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CD'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CE'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CF'.$row)->getValue()."'";
           $sql .= ",'".$itemUnidade."'";
           $sql .= ",".$worksheet->getCell('CH'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CI'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CJ'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('CK'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CL'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('CM'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CN'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CO'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CP'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CQ'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('CR'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('CS'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('CT'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CU'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CV'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('CW'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('CX'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CY'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('CZ'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('DA'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('DB'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DC'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('DD'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('DE'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DF'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('DG'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('DH'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DI'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('DJ'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DK'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DL'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DM'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DN'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DO'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DP'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DQ'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DR'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DS'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DT'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DU'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DV'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DW'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('DX'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('DY'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('DZ'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('EA'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('EB'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('EC'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('ED'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('EE'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('EF'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('EG'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('EH'.$row)->getValue();
           $sql .= ",".$worksheet->getCell('EI'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('EJ'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('EK'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('EL'.$row)->getValue()."'";
           $sql .= ",".$worksheet->getCell('EM'.$row)->getValue();
           $sql .= ",'".$worksheet->getCell('EN'.$row)->getValue()."'";
           $sql .= ",'".$worksheet->getCell('EO'.$row)->getValue()."'";
           $sql .= ")";

       $adicionarNull = explode(',', $sql);
       for ($i=0; $i < count($adicionarNull)-1 ; $i++) { 
            if (empty($adicionarNull[$i]) or $adicionarNull[$i] == "''" ) {
             $adicionarNull[$i] = "null";
            }
         }

         $sql = implode(',', $adicionarNull);

        try{
          $stmt = $database->prepare($sql);
          if ($stmt->execute()){
            close_database($database);
            $_SESSION['msg'] = '<div class="alert alert-success alert-dismissible fade show" role="alert">
                  A planilha foi <strong>processada</strong> e <strong>cadastrada</strong> com sucesso!
                  <button type="button" class="close" data-dismiss="alert" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                  </button>
                </div>';
        }else{
             $_SESSION['msg'] = '<div class="alert alert-danger alert-dismissible fade show" role="alert">
                  <strong>Erro</strong> problemas ao cadastrar processar e cadastrar planilha!
                  <button type="button" class="close" data-dismiss="alert" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                  </button>
                </div>';
        }
        }catch (Exception $e) {

         $_SESSION['msg'] = '<div class="alert alert-danger alert-dismissible fade show" role="alert">
                  <strong>Erro!</strong>
                  <button type="button" class="close" data-dismiss="alert" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                  </button>
                </div>';
      }

}
