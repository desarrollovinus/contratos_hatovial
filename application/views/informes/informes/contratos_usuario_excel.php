<?php
//Se crea un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

//Se establece la configuracion general
$objPHPExcel->getProperties()
    ->setCreator("John Arley Cano Salinas - Hatovial S.A.S.")
    ->setLastModifiedBy("John Arley Cano Salinas")
    ->setTitle("Sistema de Gestión de Contratos - Generado el ".$this->auditoria_model->formato_fecha(date('Y-m-d')).' - '.date('h:i A'))
    ->setSubject("Consolidado de hojas de vida")
    ->setDescription("Contratos por usuario")
    ->setKeywords("Contratos usuario")
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
// $objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(2);

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
    'borders' => array(
        'top' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
        ),
        'left' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
        ),
        'bottom' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
        ),
        'right' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
        ),
    ),
);

/*
 * Definicion de la anchura de las columnas
 */
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(4);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(9);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(35);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(14);

/**
 * Definición de altura de las filas
 */
// $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(15);

//Celdas a combinar
$objPHPExcel->getActiveSheet()->mergeCells('A1:N1');
$objPHPExcel->getActiveSheet()->mergeCells('A2:A3');
$objPHPExcel->getActiveSheet()->mergeCells('B2:B3');
$objPHPExcel->getActiveSheet()->mergeCells('C2:C3');
$objPHPExcel->getActiveSheet()->mergeCells('C2:C3');
$objPHPExcel->getActiveSheet()->mergeCells('D2:D3');
$objPHPExcel->getActiveSheet()->mergeCells('E2:E3');
$objPHPExcel->getActiveSheet()->mergeCells('F2:F3');
$objPHPExcel->getActiveSheet()->mergeCells('G2:G3');
$objPHPExcel->getActiveSheet()->mergeCells('H2:H3');
$objPHPExcel->getActiveSheet()->mergeCells('I2:J2');
$objPHPExcel->getActiveSheet()->mergeCells('K2:K3');
$objPHPExcel->getActiveSheet()->mergeCells('L2:L3');
$objPHPExcel->getActiveSheet()->mergeCells('M2:M3');
$objPHPExcel->getActiveSheet()->mergeCells('N2:N3');

if ($nombre_usuario) {
    $nombre = 'Contratos creados por '.$nombre_usuario;
}else{
    $nombre = 'Listado general de contratos';
}

//
$fila = 2;

//
$objPHPExcel->getActiveSheet()->setCellValue("I{$fila}", 'Adiciones');

// 
$fila++;

//Encabezados
$objPHPExcel->getActiveSheet()->setCellValue('A1', $nombre);
$objPHPExcel->getActiveSheet()->setCellValue('A2', '#');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'Numero');
$objPHPExcel->getActiveSheet()->setCellValue('C2', 'Objeto del contrato');
$objPHPExcel->getActiveSheet()->setCellValue('D2', 'Contratista');
$objPHPExcel->getActiveSheet()->setCellValue('E2', 'Fecha de inicio');
$objPHPExcel->getActiveSheet()->setCellValue('F2', 'Fecha de vencimiento');
$objPHPExcel->getActiveSheet()->setCellValue('G2', 'Plazo (días)');
$objPHPExcel->getActiveSheet()->setCellValue('H2', 'Valor inicial');
$objPHPExcel->getActiveSheet()->setCellValue('I2', 'Plazo (días)');
$objPHPExcel->getActiveSheet()->setCellValue('J3', 'Valor');
$objPHPExcel->getActiveSheet()->setCellValue('K2', 'Valor total');
$objPHPExcel->getActiveSheet()->setCellValue('L2', 'Valor pagado');
$objPHPExcel->getActiveSheet()->setCellValue('M2', 'Saldo');
$objPHPExcel->getActiveSheet()->setCellValue('N2', 'Estado');

//Fila para iniciar el contenido
$fila++;

//Contador
$cont = 1;

foreach ($contratos as $contrato) {
    //Contenido
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.$fila, number_format($cont, 0, '', '.'))
        ->setCellValue('B'.$fila, $contrato->Numero)
        ->setCellValue('C'.$fila, $contrato->Objeto)
        ->setCellValue('D'.$fila, $contrato->Contratista)
        ->setCellValue('E'.$fila, $contrato->Fecha_Inicial)
        ->setCellValue('F'.$fila, $contrato->Fecha_Vencimiento)
        ->setCellValue('H'.$fila, $contrato->Valor_Inicial);

    $objPHPExcel->getActiveSheet()->setCellValue("G{$fila}", number_format($contrato->Plazo_Inicial, 0, '', '.'));
    $objPHPExcel->getActiveSheet()->setCellValue("I{$fila}", $contrato->Plazo_Adiciones);
    $objPHPExcel->getActiveSheet()->setCellValue("J{$fila}", "$ ".number_format($contrato->Valor_Adiciones), 0, '', '.');
    $objPHPExcel->getActiveSheet()->setCellValue("K{$fila}", "$ ".number_format($contrato->Valor_Adiciones + $contrato->Valor_Inicial), 0, '', '.');
    $objPHPExcel->getActiveSheet()->setCellValue("L{$fila}", "$ ".number_format($contrato->Pagado), 0, '', '.');
    $objPHPExcel->getActiveSheet()->setCellValue("M{$fila}", "$ ".number_format(($contrato->Valor_Adiciones + $contrato->Valor_Inicial - $contrato->Pagado), 0, '', '.'));
    $objPHPExcel->getActiveSheet()->setCellValue("N{$fila}", $contrato->Estado);


    //Se definen los estilos para las celdas que llevarán el contenido
    $objPHPExcel->getActiveSheet()->getStyle("H{$fila}")->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);


    //Aumentamos la fila y contador
    $fila++;
    $cont++;
} // Foreach

/**
 * Aplicacion de los estilos
 */
$objPHPExcel->getActiveSheet()->getStyle('A1:N2')->applyFromArray($titulo_centrado_negrita);
$objPHPExcel->getActiveSheet()->getStyle('I3:J3')->applyFromArray($titulo_centrado_negrita);








//Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
// header("Content-Type:   application/vnd.ms-excel; charset=utf-8");
header('Cache-Control: max-age=0');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="Hojas de vida.xlsx"');

//Se genera el excel
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>