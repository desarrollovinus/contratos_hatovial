<?php
//Se crea un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

//Se establece la configuracion general
$objPHPExcel->getProperties()
    ->setCreator("John Arley Cano Salinas - Hatovial S.A.S.")
    ->setLastModifiedBy("John Arley Cano Salinas")
    ->setTitle("Sistema de Gestión de Contratos - Generado el ".$this->auditoria_model->formato_fecha(date('Y-m-d')).' - '.date('h:i A'))
    ->setSubject("Contratos categorizados por subcontratista")
    ->setDescription("Contratos por subcontratista")
    ->setKeywords("Contratos subcontratista")
    ->setCategory("Reporte");

//Definicion de las configuraciones por defecto en todo el libro
$objPHPExcel->getDefaultStyle()->getFont()->setName('Helvetica'); //Tipo de letra
$objPHPExcel->getDefaultStyle()->getFont()->setSize(8); //Tamanio
$objPHPExcel->getDefaultStyle()->getAlignment()->setWrapText(true);//Ajuste de texto
$objPHPExcel->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);// Alineacion centrada

//Se establece la configuracion de la pagina
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE); //Orientacion horizontal
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_LEGAL); //Tamano carta
$objPHPExcel->getActiveSheet()->getPageSetup()->setScale(100); //Escala

//Se indica el rango de filas que se van a repetir en el momento de imprimir. (Encabezado del reporte)
$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(2);

//Se establecen las margenes
$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0,90); //Arriba
$objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0,70); //Derecha
$objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0,70); //Izquierda
// $objPHPExcel->getActiveSheet()->getPageMargins()->setBottom(0,500); //Abajo

//Centrar página
$objPHPExcel->getActiveSheet()->getPageSetup()->setHorizontalCentered();

/**
 * Estilos
 */
//Estilo de los titulos
$titulo_centrado_negrita = array(
    'font' => array(
        'bold' => true
    ),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER
    )
);

$titulo_centrado = array(
    'font' => array(
        'bold' => false
    ),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER
    )
);

$titulo_derecho = array(
    'font' => array(
        'bold' => true
    ),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
    )
);

//Estilo de los bordes
$bordes = array(
    'borders' => array(
        'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array(
                'argb' => '000000'
            )
        ),
    ),
);

//Array de estilos para los n&uacute;meros 
$estilos_numeros = array(
    'font' => array(
        'bold' => false,
    ),
    'alignment' => array(
        //Ajustando el textos (Derecha, centro, izquierda)
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    ),
);

/*
 * Definicion de la anchura de las columnas
 */
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(4);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(40);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(11);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(7);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(17);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(17);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(17);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(17);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(17);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(10);

/**
 * Definición de altura de las filas
 */
$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(20);

//Celdas a combinar
$objPHPExcel->getActiveSheet()->mergeCells('A1:O1');

//Encabezados
$objPHPExcel->getActiveSheet()->setCellValue('A1', "CONTRATOS ASOCIADOS A " . $nombre);
$objPHPExcel->getActiveSheet()->setCellValue('A2', '#');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'Número');
$objPHPExcel->getActiveSheet()->setCellValue('C2', 'Objeto del contrato');
$objPHPExcel->getActiveSheet()->setCellValue('D2', 'Contratista');
$objPHPExcel->getActiveSheet()->setCellValue('E2', 'Contratante');
$objPHPExcel->getActiveSheet()->setCellValue('F2', 'Fecha de inicio');
$objPHPExcel->getActiveSheet()->setCellValue('G2', 'Fecha de vencimiento');
$objPHPExcel->getActiveSheet()->setCellValue('H2', 'Plazo (días)');
$objPHPExcel->getActiveSheet()->setCellValue('I2', 'Valor inicial');
$objPHPExcel->getActiveSheet()->setCellValue('J2', 'Plazo adiciones');
$objPHPExcel->getActiveSheet()->setCellValue('K2', 'Valor adiciones');
$objPHPExcel->getActiveSheet()->setCellValue('L2', 'Valor total');
$objPHPExcel->getActiveSheet()->setCellValue('M2', 'Valor pagado');
$objPHPExcel->getActiveSheet()->setCellValue('N2', 'Saldo');
$objPHPExcel->getActiveSheet()->setCellValue('O2', 'Estado');

/**
 * Estilos de los encabezados
 */
$objPHPExcel->getActiveSheet()->getStyle('A1:O2')->applyFromArray($titulo_centrado_negrita);
$objPHPExcel->getActiveSheet()->getStyle('A2:O2')->applyFromArray($bordes);

//Fila para iniciar el contenido
$fila = 3;
$fila_inicial = $fila;

//Contador
$cont = 1;

foreach ($contratos as $contrato) {
    //Contenido
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.$fila, number_format($cont, 0, '', '.'))
        ->setCellValue('B'.$fila, $contrato->Numero)
        ->setCellValue('C'.$fila, $contrato->Objeto)
        ->setCellValue('D'.$fila, $contrato->Contratista)
        ->setCellValue('E'.$fila, $contrato->Contratante)
        ->setCellValue('F'.$fila, $contrato->Fecha_Inicial)
        ->setCellValue('G'.$fila, $contrato->Fecha_Vencimiento)
        ->setCellValue('H'.$fila, $contrato->Plazo_Inicial)
        ->setCellValue('I'.$fila, $contrato->Valor_Inicial)
        ->setCellValue('J'.$fila, $contrato->Plazo_Adiciones)
        ->setCellValue('K'.$fila, $contrato->Valor_Adiciones)
        ->setCellValue('L'.$fila, $contrato->Pagado)
        ->setCellValue('M'.$fila, $contrato->Valor_Adiciones + $contrato->Valor_Inicial)
        ->setCellValue('N'.$fila, ($contrato->Valor_Adiciones + $contrato->Valor_Inicial) - $contrato->Pagado)
        ->setCellValue('O'.$fila, $contrato->Estado);

    //Se definen los estilos para las celdas que llevarán el contenido
    
    $objPHPExcel->getActiveSheet()->getStyle("I{$fila}")->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
    $objPHPExcel->getActiveSheet()->getStyle("K{$fila}")->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
    $objPHPExcel->getActiveSheet()->getStyle("L{$fila}")->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
    $objPHPExcel->getActiveSheet()->getStyle("M{$fila}")->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
    $objPHPExcel->getActiveSheet()->getStyle("N{$fila}")->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);

    //Aumentamos la fila y contador
    $fila++;
    $cont++;
} // foreach

//Se disminuye una fila
$fila--;

/**
 * Aplicacion de los estilos
 */
$objPHPExcel->getActiveSheet()->getStyle("A{$fila_inicial}:O{$fila}")->applyFromArray($bordes);

//Pié de página
$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' .$objPHPExcel->getProperties()->getTitle() . '&RPágina &P de &N');

//Cambiar espacios por underline
$nombre = str_replace(' ', '_', $nombre);

//Título de la hoja
$objPHPExcel->getActiveSheet()->setTitle($nombre);

//Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
header('Cache-Control: max-age=0');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename='Contratos_{$nombre}.xlsx'");

//Se genera el excel
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>