<?php
// Se obtiene el número de contrato
$id_contrato = $this->uri->segment(3);

// Se consultan los datos del contrato
$contratos = $this->contrato_model->listar_contratos($id_contrato);

//Se traen los pagos existentes de ese contrato
$pagos = $this->pago_model->listar_pagos($id_contrato);

// Se crea el nuevo objeto
$PHPWord = new PHPWord();

/**
 * Configuración por defecto
 */
$PHPWord->setDefaultFontName('Arial');
$PHPWord->setDefaultFontSize(11);
$properties = $PHPWord->getProperties();
$properties->setCreator('John Arley Cano Salinas'); 
$properties->setCompany('Hatovial S.A.S.');
$properties->setTitle('Acta de liquidación de obra');
$properties->setDescription('Acta de liquidación de obra'); 
$properties->setCategory('informe');
$properties->setLastModifiedBy('John Arley Cano Salinas');

/**
 * Estilos de alineación
 */
$alineacion_centrada = array ('align' => 'center', 'valign' => 'center');
$alineacion_izquierda = array ('align' => 'left', 'valign' => 'center');
$alineacion_derecha = array ('align' => 'right', 'valign' => 'center');
$alineacion_justificada = array ('align' => 'both', 'valign' => 'center');

/**
 * Estilos de las fuentes
 */
$PHPWord->addFontStyle('titulo1', array( 'name'=>'Arial', 'size'=> 18, 'color'=>'000000', 'bold'=> true, 'italic' => true));
$PHPWord->addFontStyle('titulo2', array( 'name'=>'Arial', 'size'=> 12, 'color'=>'000000', 'bold'=> true));
$PHPWord->addFontStyle('titulo3', array( 'name'=>'Arial', 'size'=> 8, 'color'=>'000000', 'bold'=> true, 'italic' => true));

$PHPWord->addFontStyle('parrafo1', array( 'name'=>'Arial', 'size'=> 11, 'color'=>'000000', 'bold'=> true));
$PHPWord->addFontStyle('parrafo2', array( 'name'=>'Arial', 'size'=> 12, 'color'=>'000000', 'bold'=> false));
$PHPWord->addFontStyle('parrafo2', array( 'name'=>'Arial', 'size'=> 8, 'color'=>'000000', 'bold'=> false));

/**
 * Estilos de las tablas
 */
$tabla1 =array('borderColor'=>'000000', 'borderSize'=> 6);
$tabla2 = array('borderSize' => 8, 'borderColor' => '1E1E1E',  'cellMarginTop' => 100, 'rules' => 'cols');
$tabla3 = array('cellMarginTop' => 50, 'rules' => 'cols');
$tabla4 = array(    'borderRightSize' => 50, 'borderBottomColor' => '009900',    'borderBottomSize' => 50, 'borderRightColor' => '00CCFF',    'borderTopSize' => 50, 'borderTopColor' => '00CCFF',    'borderLeftSize' => 50, 'borderLeftColor' => '00CCFF'); 

/**
 * Estilos de celdas
 */
$styleCell = array('valign' => 'center');

/**
 * Sección 1
 */
$seccion1 = $PHPWord->createSection(array(/* 'marginLeft'=>210, 'marginRight'=>200, 'marginTop'=>620, 'marginBottom'=>0,*/'pageSizeW'=>12240, 'pageSizeH'=>15840));

/**
 * Cabecera
 */
$PHPWord->addTableStyle('tabla1', $tabla1);
$cabecera = $seccion1->createHeader();
$table = $cabecera->addTable('tabla1');
$table->addRow(1000);
// $table->addCell(4000, $styleCell)->addText('COCAN 900.193.471-9', 'titulo1', $alineacion_centrada);
$table->addCell(10000, $styleCell)->addText(utf8_decode("ACTA DE LIQUIDACIÓN DE OBRA"), 'titulo2', $alineacion_centrada);
$cabecera->addTextBreak();

/**
 * Pié de página
 */
$footer = $seccion1->createFooter();
$footer->addPreserveText(utf8_decode('Calle 59 No. 48-35 Autopista Norte, Copacabana - Página {PAGE} de {NUMPAGES}'), 'titulo3', $alineacion_centrada);

$texto1 = file_get_contents('application/views/informes/plantilla1.txt');
$texto2 = file_get_contents('application/views/informes/plantilla2.txt');
$texto3 = file_get_contents('application/views/informes/plantilla3.txt');

foreach ($contratos as $contrato):
	$texto1 = file_get_contents('application/views/informes/plantilla1.txt');
    $texto1 = str_replace('{CONTRATO}', $contrato->Numero, $texto1);
    $texto1 = str_replace('{CONTRATISTA}', $contrato->Contratista, $texto1);
    $texto1 = str_replace('{DOCUMENTO_CONTRATISTA}', $contrato->Documento_Contratista, $texto1);
    $texto1 = str_replace('{REPRESENTANTE_LEGAL}', strtoupper($contrato->Representante_Legal), $texto1);

    // Texto principal
	$seccion1->addText(utf8_decode($texto1), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();

	// Objeto	
	$seccion1->addText('OBJETO DEL CONTRATO', 'parrafo1', $alineacion_izquierda);
	$seccion1->addText(utf8_decode($contrato->Objeto), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();

	// Valor	
    $numero = new NumerosALetras();
	$seccion1->addText('VALOR DEL CONTRATO', 'parrafo1', $alineacion_izquierda);
	$seccion1->addText(utf8_decode('$ '.number_format($contrato->Valor_Inicial,2,',','.').' ('.$numero->traducir($contrato->Valor_Inicial).')'), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();

	// Fecha de inicio
	$seccion1->addText(utf8_decode('FECHA DE INICIACIÓN: '.$this->auditoria_model->formato_fecha($contrato->Fecha_Inicial)), 'parrafo1', $alineacion_izquierda);

	// Fecha final
	$seccion1->addText(utf8_decode('FECHA DE VENCIMIENTO: '.$this->auditoria_model->formato_fecha($contrato->Fecha_Vencimiento)), 'parrafo1', $alineacion_izquierda);
	$seccion1->addTextBreak();

	// Valor total
	$seccion1->addText('VALOR TOTAL EJECUTADO POR EL CONTRATISTA', 'parrafo1', $alineacion_izquierda);
	$seccion1->addText(utf8_decode('$ '.number_format($contrato->Valor_Inicial,2,',','.').' ('.$numero->traducir($contrato->Valor_Inicial).')'), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();

	$texto2 = str_replace('{FECHA_VENCIMIENTO}', $this->auditoria_model->formato_fecha($contrato->Fecha_Vencimiento), $texto2);
    $texto2 = str_replace('{REPRESENTANTE_LEGAL}', strtoupper($contrato->Representante_Legal), $texto2);
    $texto2 = str_replace('{CONTRATISTA}', $contrato->Contratista, $texto2);
	$seccion1->addText(utf8_decode($texto2), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();

	$numero = 1;
    foreach ($pagos as $pago):
    	if($numero == 1){
			$seccion1->addText('BALANCE GENERAL DEL CONTRATO', 'parrafo1', $alineacion_izquierda);
			$seccion1->addText(utf8_decode('A continuación se presentará el balance financiero de la oferta'), 'parrafo2', $alineacion_izquierda);
        }

        $table = $seccion1->addTable();
		$table->addRow();
		$table->addCell(10000)->addText(utf8_decode('VALOR FACTURADO EN ACTA '.$pago->Numero_Acta.' ('.$this->auditoria_model->formato_fecha($pago->Fecha).')'), 'parrafo2', $alineacion_izquierda);
		$table->addCell(5000)->addText(utf8_decode('$ '.number_format($pago->Valor_Pago, 0, '', '.')), 'parrafo2', $alineacion_derecha);
		
		$table->addRow();
		$table->addCell(10000)->addText(utf8_decode('Rete garantía (5% antes de IVA)'), 'parrafo2', $alineacion_izquierda);
		$table->addCell(5000)->addText(utf8_decode('$ '.number_format($pago->Valor_Retenido, 0, '', '.')), 'parrafo2', $alineacion_derecha);
		
		$seccion1->addText((utf8_decode('FACTURA '.$numero.' ('.$contrato->Contratista.')')), 'parrafo2', $alineacion_izquierda);

		$seccion1->addTextBreak();
        $numero++;
    endforeach;

    $numero = 1;
    foreach ($pagos as $pago):
        if($numero == 1){
			$seccion1->addText( utf8_decode('RELACIÓN DE RETENIDOS'), 'parrafo1', $alineacion_izquierda);
        }

		$seccion1->addText(utf8_decode('Acta Factura '.$numero.': Rete garantía (5% antes del IVA) $ '.number_format($pago->Valor_Retenido, 0, '', '.')), 'parrafo2', $alineacion_justificada);

        $numero++;
    endforeach;
	$seccion1->addTextBreak();

	$seccion1->addText( utf8_decode('DOCUMENTOS APORTADOS EN LA LIQUIDACIÓN'), 'parrafo1', $alineacion_izquierda);
	$seccion1->addText( utf8_decode($texto3), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();

	$seccion1->addText( utf8_decode('Para constancia se firma por quienes en ella intervinieron.'), 'parrafo2', $alineacion_justificada);
	$seccion1->addTextBreak();
	$seccion1->addTextBreak();

	$table = $seccion1->addTable();
	$table->addRow();
	$table->addCell(9000)->addText(utf8_decode('Ing. RICARDO LÓPEZ LOMBANA'), 'parrafo1', $alineacion_izquierda);
	$table->addCell(9000)->addText(utf8_decode('Ing. '.strtoupper($contrato->Representante_Legal)), 'parrafo1', $alineacion_izquierda);
	
	$table->addRow();
	$table->addCell(9000)->addText(utf8_decode('Representante legal'), 'parrafo2', $alineacion_izquierda);
	$table->addCell(9000)->addText(utf8_decode('Representante legal'), 'parrafo2', $alineacion_izquierda);

	$table->addRow();
	$table->addCell(9000)->addText(utf8_decode('Consorcio Constructor Aburrá Norte'), 'parrafo2', $alineacion_izquierda);
	$table->addCell(9000)->addText(utf8_decode($contrato->Contratista), 'parrafo2', $alineacion_izquierda);

	$table->addRow();
	$table->addCell(9000)->addText(utf8_decode('COCAN'), 'parrafo2', $alineacion_izquierda);
endforeach;

// At least write the document to webspace:
$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');

$temp_file_uri = tempnam('', 'xyz');
$objWriter->save($temp_file_uri);

//download code
header('Content-Description: File Transfer');
header('Content-Type: application/octet-stream');
// header('Content-Disposition: attachment; filename=Acta_Recibo.docx');
header('Content-Disposition: attachment; filename=Acta_Liquidacion_'.$contrato->Numero.'.docx');
header('Content-Transfer-Encoding: binary');
header('Expires: 0');
header('Content-Length: ' . filesize($temp_file_uri));
readfile($temp_file_uri);
unlink($temp_file_uri); // deletes the temporary file
exit;
?>