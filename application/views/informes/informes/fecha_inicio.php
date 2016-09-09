<?php
//Se reciben las fechas
$fecha = 'Fecha_Inicial';
$fecha1 = $this->input->post('fecha1');
$fecha2 = $this->input->post('fecha2');

//Se carga el modelo que va a traer todos los datos
$contratos = $this->informes_model->contratos_por_fecha_inicial($fecha, $fecha1, $fecha2);

//Se crea un nuevo objeto PHPExcel
$objPHPExcel = new PHPExcel();

//Se establecen las propiedades del documento
$objPHPExcel->getProperties()
    ->setCreator('John Arley Cano - Hatovial S.A.S.')
    ->setLastModifiedBy('John Arley Cano - Hatovial S.A.S.')
    ->setTitle('Contratos por fecha inicial')
    ->setSubject("Contratos filtrados por la fecha inicial")
    ->setDescription("Listado de contratos existentes a la fecha seg&uacute;n el rango de fechas seleccionado")
    ->setKeywords("fecha inicio contratos excel hatovial")
    ->setCategory("Informe");

//Se establece el tamaño y fuente del texto
$objPHPExcel->getDefaultStyle()->getFont()->setName('Tahoma');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(10);

/*
* Se establece el tamaño de las columnas (en pixeles)
* Si se quiere automáticamente el tamaño: getColumnDimension('B')->setAutoSize(true)
*/
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(45);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);

/*
* Aqu&iacute; se establecen las celdas que se van a combinar
*/
$objPHPExcel->getActiveSheet()->mergeCells('A1:H1');

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
$objPHPExcel->getActiveSheet()->getStyle('A1:H1')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('A2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('B2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('C2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('D2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('E2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('F2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('G2')->applyFromArray($estilos_titulos);
$objPHPExcel->getActiveSheet()->getStyle('H2')->applyFromArray($estilos_titulos);

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
        ->setCellValue('C'.$fila, $contrato->Lugar)
        ->setCellValue('D'.$fila, $contrato->Valor_Inicial)
        ->setCellValue('E'.$fila, $contrato->Fecha_Inicial)
        ->setCellValue('F'.$fila, $contrato->Fecha_Vencimiento)
        ->setCellValue('G'.$fila, $contrato->Contratista)
        ->setCellValue('H'.$fila, $contrato->Estado);

    //Se definen los estilos para las celdas que llevarán el contenido
    $objPHPExcel->getActiveSheet()->getStyle('A'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('B'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('C'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('D'.$fila)->applyFromArray($estilos_numeros)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
    $objPHPExcel->getActiveSheet()->getStyle('E'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('F'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('G'.$fila)->applyFromArray($estilos_textos);
    $objPHPExcel->getActiveSheet()->getStyle('H'.$fila)->applyFromArray($estilos_textos);
        
    $fila++;
endforeach;

/*
 * Encabezados del reporte
 */
$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', 'Contratos con fecha de inicio desde '.$fecha1.' hasta '.$fecha2)
    ->setCellValue('A2', 'Número')
    ->setCellValue('B2', 'Objeto')
    ->setCellValue('C2', 'Localización')
    ->setCellValue('D2', 'Valor Inicial')
    ->setCellValue('E2', 'Fecha Inicial')
    ->setCellValue('F2', 'Fecha Vencimiento')
    ->setCellValue('G2', 'Contratista')
    ->setCellValue('H2', 'Estado');

/*
* Configuraciones finales antes de generar el archivo
*/
//Se establece la orientación de la hoja
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);

//Se establece el tipo de papel
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_LETTER);

//Se establece el nombre de la hoja
$objPHPExcel->getActiveSheet()->setTitle('Contratos');

// Se establece la hoja activa, para que cuando se abra el documento se muestre primero.
$objPHPExcel->setActiveSheetIndex(0);

//Se indica el rango de filas que se van a repetir en el momento de imprimir. (Encabezado del reporte)
$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(2);

//Se ajusta el área de impresión al contenido, definiendo un 80%
$objPHPExcel->getActiveSheet()->getPageSetup()->setScale(80)->setFitToPage(false);

if($fila == 3){
    //Se establece el mensaje de error
    $this->data['mensaje_alerta'] = 'No se ha generado ning&uacute;n informe porque no hay<br/>contratos en el rango de fechas seleccionado';

    //Se traen los estados de los contratos
    $this->data['contratos_estados'] = $this->contrato_model->listar_contratos_estados();
    //se establece el titulo de la p&aacute;gina
    $this->data['titulo'] = 'Informes';
    //se establece la vista que tiene el contenido principal
    $this->data['contenido_principal'] = 'informes/informes_view';
    //se carga el template
    $this->load->view('includes/template', $this->data);
}else{
    //Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
    header('Cache-Control: max-age=0');
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="Contratos desde '.$fecha1.' hasta '.$fecha2.'.xlsx"');

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');
    exit;
}
?>