<?php
//Se crea un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

//Se establece la configuracion general
$objPHPExcel->getProperties()
    ->setCreator("Luis David Moreno Lopera - Hatovial S.A.S.")
    ->setLastModifiedBy("Luis David Moreno Lopera")
    ->setTitle("Sistema de Gestión de Contratos - Generado el ".$this->auditoria_model->formato_fecha(date('Y-m-d')).' - '.date('h:i A'))
    ->setSubject("Contratos categorizados por subcontratista")
    ->setDescription("Pagos a contrato")
    ->setKeywords("Pagos contrato")
    ->setCategory("Reporte");

//Definicion de las configuraciones por defecto en todo el libro
$objPHPExcel->getDefaultStyle()->getFont()->setName('ARIAL'); //Tipo de letra
$objPHPExcel->getDefaultStyle()->getAlignment()->setWrapText(true);//Ajuste de texto
$objPHPExcel->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);// Alineacion centrada

//Se establece la configuracion de la pagina
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE); //Orientacion horizontal
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_LETTER); //Tamano carta
$objPHPExcel->getActiveSheet()->getPageSetup()->setScale(100); //Escala

// Se establecen las margenes
$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0,90); //Arriba
$objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0,70); //Derecha
$objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0,70); //Izquierda

$hoja = $objPHPExcel->getActiveSheet();
$hoja->setTitle($estado_pagos[0]->Numero);
//Se indica el rango de filas que se van a repetir en el momento de imprimir. (Encabezado del reporte)
$hoja->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(3);
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

$alineacion_izquierda = array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
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

$derecha_negrita = array(
    'font' => array(
        'bold' => true,
    ),
    'alignment' => array(
        //Ajustando el textos (Derecha, centro, izquierda)
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    ),
);


// asignacion de tamaño de la fila 1
$hoja->getRowDimension(1)->setRowHeight(30);

// asignacion de tamaño para las filas
for ($i=3; $i <= 50; $i++) {
	$hoja->getRowDimension($i)->setRowHeight(20);
}

 // asignacion de tamaño para las columnas
for ($columna="A"; $columna <= "D"; $columna++) {
	$hoja->getColumnDimension($columna)->setWidth(21);
}

$fila = 1;
$hoja->mergeCells("A{$fila}:D{$fila}");
$objPHPExcel->getDefaultStyle()->getFont()->setSize(14); //Tamaño fuente
$hoja->getStyle("A{$fila}:D{$fila}")->applyFromArray($titulo_centrado_negrita);
$hoja->setCellValue("A{$fila}", 'Pagos a contrato '.$estado_pagos[0]->Numero);
$fila++;

$hoja->getRowDimension($fila)->setRowHeight(7);
$fila++;

$objPHPExcel->getDefaultStyle()->getFont()->setSize(11); //Tamaño fuente
$hoja->getRowDimension($fila)->setRowHeight(25);
$hoja->getStyle("A{$fila}:D{$fila}")->applyFromArray($bordes);
$hoja->getStyle("A{$fila}:D{$fila}")->applyFromArray($titulo_centrado_negrita);
$hoja->setCellValue("A{$fila}", 'Fecha');
$hoja->setCellValue("B{$fila}", 'Número de acta');
$hoja->setCellValue("C{$fila}", 'Valor');
$hoja->setCellValue("D{$fila}", 'Retenido');
$fila++;

$fila_inicio = $fila;
foreach ($pagos as $pago) {
    $objPHPExcel->getActiveSheet()->getStyle("C{$fila}:D{$fila}")
    ->getNumberFormat()
    ->setFormatCode('$ * ###,###,###,###');
    $hoja->getStyle("A{$fila}:D{$fila}")->applyFromArray($bordes);
    $hoja->getStyle("B{$fila}")->applyFromArray($alineacion_izquierda);
    $hoja->setCellValue("A{$fila}", substr($pago->Fecha, 8, 2)."-".substr($pago->Fecha, 5, 2)."-".substr($pago->Fecha, 0, 4));
    $hoja->setCellValue("B{$fila}", $pago->Numero_Acta);
    $hoja->setCellValue("C{$fila}", $pago->Valor_Pago);
    $hoja->setCellValue("D{$fila}", $pago->Valor_Retenido);
    $fila++;
}
$fila_fin = $fila - 1;

$hoja->getRowDimension($fila)->setRowHeight(7);
$fila++;

$fila_aux = $fila + 3;
$objPHPExcel->getActiveSheet()->getStyle("B{$fila}:D{$fila_aux}")
->getNumberFormat()
->setFormatCode('$ * ###,###,###,###');
$hoja->getStyle("A{$fila}:D{$fila_aux}")->applyFromArray($derecha_negrita);

$valor = $estado_pagos[0]->Valor_Inicial + $estado_pagos[0]->Valor_Adiciones;
$hoja->mergeCells("B{$fila}:C{$fila}");
$hoja->setCellValue("B{$fila}", "Valor (incluidas adiciones)");
$hoja->setCellValue("D{$fila}", $valor);
$fila_Valor_Inicial = $fila;
$fila++;


$hoja->mergeCells("B{$fila}:C{$fila}");
$hoja->setCellValue("B{$fila}", "Total pagado");
$hoja->setCellValue("D{$fila}", "=SUM(C{$fila_inicio}:C{$fila_fin})");
$fila_Pagado = $fila;
$fila++;


$hoja->mergeCells("B{$fila}:C{$fila}");
$hoja->setCellValue("B{$fila}", "Saldo");
$hoja->setCellValue("D{$fila}", "=(D{$fila_Valor_Inicial}-D{$fila_Pagado})");
$fila++;

$hoja->mergeCells("B{$fila}:C{$fila}");
$hoja->setCellValue("B{$fila}", "Total retenido");
$hoja->setCellValue("D{$fila}", "=SUM(D{$fila_inicio}:D{$fila_fin})");
$fila++;

$hoja->getStyle("D{$fila}")
    ->getNumberFormat()->applyFromArray(
        array(
            'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00
        )
    );
$hoja->getStyle("A{$fila}:D{$fila}")->applyFromArray($derecha_negrita);
$hoja->mergeCells("B{$fila}:C{$fila}");
$hoja->setCellValue("B{$fila}", "Porcentaje");
$hoja->setCellValue("D{$fila}", "=(D{$fila_Pagado}/D{$fila_Valor_Inicial})");

//
// Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel. SUMA(D12;-D13)
header('Cache-Control: max-age=0');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename='Pagos a ".$estado_pagos[0]->Numero.".xlsx'");

//Se genera el excel
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>
