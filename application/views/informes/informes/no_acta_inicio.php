<?php
//Se carga el modelo
$contratos = $this->email_model->no_acta_inicio();

//Se crea un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

//Se establecen las propiedades del documento
$objPHPExcel->getProperties()
    ->setCreator('John Arley Cano - Hatovial S.A.S.')
    ->setLastModifiedBy('John Arley Cano - Hatovial S.A.S.')
    ->setTitle('Contratos sin acta de inicio')
    ->setSubject("Contratos que no tienen acta de inicio")
    ->setDescription("Listado de contratos que a la fecha no tienen acta de inicio")
    ->setKeywords("acta inicio contratos excel hatovial")
    ->setCategory("Informe");

//Se establece el tamaño y fuente del texto
$objPHPExcel->getDefaultStyle()->getFont()->setName('Tahoma');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(10);

/*
* Se establece el tamaño de las columnas (en pixeles)
* Si se quiere automáticamente el tamaño: getColumnDimension('B')->setAutoSize(true)
*/
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(64);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(17);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);

/*
* Aqu&iacute; se establecen las celdas que se van a combinar
*/
$objPHPExcel->getActiveSheet()->mergeCells('A1:G1');

/*
* Estableceremos arrays para cada estilo de las celdas. Por ejemplo
* Para las celdas de t&iacute;tulo, para las celdas de contenido, etc.
*/
//Array de estilos para los t&iacute;tulos
$estilos_titulos = array(
    'font' => array(
        'bold' => true
    ),
    'alignment' => array(
        //Ajustando los textos (Derecha, centro, izquierda)
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
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
        )
    )
);

//Array de estilos para los textos 
$estilos_textos = array(
    'font' => array(
        'bold' => false
    ),
    'alignment' => array(
        //Ajustando el textos (Derecha, centro, izquierda)
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
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
 * Se especifica el array de estilos que se le va a aplicar a cada celda o rango de celdas
 */
//Array_estilos_titulos
$objPHPExcel->getActiveSheet()->getStyle('A1:G1')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('A2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('B2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('C2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('D2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('E2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('F2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('G2')->applyFromArray($estilos_titulos);

//Justificación del texto
//$objPHPExcel->getActiveSheet()->getStyle('A1:D4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

//Ajustar el texto
$objPHPExcel->getActiveSheet()->getStyle('A1:G200')->getAlignment()->setWrapText(true);

//Se define la fila. El contenido empieza a partir de la fila 3
$fila = 3;

/*
 * Contenido del reporte. Se generará dentro de un ciclo 
 */
foreach ($contratos as $contrato):
    //Contenido
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.$fila, $contrato->Numero)
        ->setCellValue('B'.$fila, $contrato->Objeto)
        ->setCellValue('C'.$fila, $contrato->Contratista)
        ->setCellValue('D'.$fila, $contrato->Fecha_Inicial)
        ->setCellValue('E'.$fila, $contrato->Fecha_Vencimiento)
        ->setCellValue('F'.$fila, $contrato->Valor_Inicial)
        ->setCellValue('G'.$fila, $contrato->Estado);
    
    //Se definen los estilos para las celdas que llevarán el contenido
    $objPHPExcel->getActiveSheet()->getStyle('A'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('B'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('C'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('D'.$fila)->applyFromArray($estilos_textos)->getNumberFormat();
    $objPHPExcel->getActiveSheet()->getStyle('E'.$fila)->applyFromArray($estilos_textos)->getNumberFormat();
    $objPHPExcel->getActiveSheet()->getStyle('F'.$fila)->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
    $objPHPExcel->getActiveSheet()->getStyle('G'.$fila)->applyFromArray($estilos_textos);
    
    
    $fila++;
    
    
endforeach;

/*
* Encabezados del reporte
*/
$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', 'Contratos sin acta de inicio')
    ->setCellValue('A2', 'Número')
    ->setCellValue('B2', 'Objeto')
    ->setCellValue('C2', 'Contratista')
    ->setCellValue('D2', 'Fecha inicial')
    ->setCellValue('E2', 'Fecha vencimiento')
    ->setCellValue('F2', 'Valor Inicial')
    ->setCellValue('G2', 'Estado');
    
/*
 * Configuraciones finales antes de generar el archivo
 */
//Se establece la orientación de la hoja
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

//Se establece el tipo de papel
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_LETTER);

//Se establece el nombre de la hoja
$objPHPExcel->getActiveSheet()->setTitle('Contratos sin acta de inicio');

// Se establece la hoja activa, para que cuando se abra el documento se muestre primero.
$objPHPExcel->setActiveSheetIndex(0);

//Se indica el rango de filas que se van a repetir en el momento de imprimir. (Encabezado del reporte)
$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(2);

//Se ajusta el área de impresión al contenido, definiendo un 80%
$objPHPExcel->getActiveSheet()->getPageSetup()->setScale(80)->setFitToPage(false);
    
//Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
header('Cache-Control: max-age=0');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="Contratos sin acta de inicio.xlsx"');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>