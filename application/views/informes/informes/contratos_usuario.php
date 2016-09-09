<?php
/*
 * Creaci&oacute;n del objeto de la clase heredada
 * Informacion general del informe
 */
$pdf = new PDF_Diag('L','mm','Legal');
$pdf->AddPage();
$pdf->SetAuthor('John Arley Cano - Hatovial S.A.S.');
$pdf->SetTitle('John Arley Cano - Hatovial S.A.S.');
$pdf->SetCreator('John Arley Cano - Hatovial S.A.S.');
$pdf->SetMargins(5, 5, 5, 5);

if ($nombre_usuario) {
    $nombre = 'Contratos creados por '.$nombre_usuario;
}else{
    $nombre = 'Listado general de contratos';
}

//Cabecera
$pdf->Image('img/logo2.png', 5, 5);
$pdf->SetFont('Arial','B',14);
$pdf->Cell(340, 7, utf8_decode(substr($nombre, 0, 75)), 0, 1, 'C');
$pdf->Ln(10);

//Titulos
$pdf->SetFont('Arial','B',10);
$pdf->SetXY(5, 25);
$pdf->Cell(7, 10, utf8_decode(''), 1, 0, 'C');
$pdf->SetXY(12, 25);
$pdf->Cell(16, 10, utf8_decode('Número'), 1, 0, 'C');
$pdf->SetXY(28, 25);
$pdf->Cell(60, 10, utf8_decode('Objeto del contrato'), 1, 0, 'C');
$pdf->SetXY(88, 25);
$pdf->Cell(38, 10, utf8_decode('Contratista'), 1, 0, 'C');
$pdf->SetXY(126, 25);
$pdf->Cell(22, 10, utf8_decode('Inicio'), 1, 0, 'C');
$pdf->SetXY(148, 25);
$pdf->Cell(22, 10, utf8_decode('Vencimiento'), 1, 0, 'C');
$pdf->SetXY(170, 25);
$pdf->Cell(13, 10, utf8_decode('Plazo'), 1, 0, 'C');
$pdf->SetXY(183, 25);
$pdf->Cell(23, 10,'Valor Inicial', 1, 0, 'C');
$pdf->SetXY(206, 25);
$pdf->Cell(36, 5, utf8_decode('Adiciones'), 1, 1, 'C');
$pdf->SetX(206);
$pdf->Cell(13, 5, utf8_decode('Plazo'), 1, 0, 'C');
$pdf->Cell(23, 5, utf8_decode('Valor'), 1, 0, 'C');
$pdf->SetXY(242, 25);
$pdf->Cell(23, 10,'Valor total', 1, 0, 'C');
$pdf->Cell(23, 10,'Pagado', 1, 0, 'C');
$pdf->Cell(23, 10,'Saldo', 1, 0, 'C');
$pdf->Cell(27, 10, utf8_decode('Estado'), 1, 0, 'C');

//Contador
$pdf->Ln();
$cont = 1;

//Contenido
$pdf->SetFont('Arial', '', 8);

foreach ($contratos as $contrato) {
    //Se establece el link del contrato
    $link = 'http://www.hatovial.com/contratos/index.php/ver/'.$contrato->Pk_Id_Contrato;

    //$pdf->Write(5, 'aqui', $link);
    $pdf->Cell(7, 5, $cont, 1, 0, 'R');
    $pdf->Cell(16, 5, $contrato->Numero, 1, 0, 'R');
    $pdf->Cell(60, 5, utf8_decode(substr($contrato->Objeto, 0, 38)), 1, 0, 'L');
    $pdf->Cell(38, 5, utf8_decode(substr($contrato->Contratista, 0, 17)), 1, 0, 'L');
    $pdf->Cell(22, 5, date('d-m-Y', strtotime($contrato->Fecha_Inicial)), 1, 0, 'L');
    $pdf->Cell(22, 5, date('d-m-Y', strtotime($contrato->Fecha_Vencimiento)), 1, 0, 'L');
    $pdf->Cell(13, 5, number_format($contrato->Plazo_Inicial, 0, '', '.'), 1, 0, 'R');
    $pdf->Cell(23, 5, '$ '.number_format($contrato->Valor_Inicial, 0, '', '.'), 1, 0, 'R');
    $pdf->Cell(13, 5, $contrato->Plazo_Adiciones, 1, 0, 'R');
    $pdf->Cell(23, 5, '$ '.number_format($contrato->Valor_Adiciones, 0, '', '.'), 1, 0, 'R');
    $pdf->Cell(23, 5, '$ '.number_format($contrato->Valor_Adiciones + $contrato->Valor_Inicial, 0, '', '.'), 1, 0, 'R');
    $pdf->Cell(23, 5, '$ '.number_format($contrato->Pagado, 0, '', '.'), 1, 0, 'R');
    $pdf->Cell(23, 5, '$ '.number_format($contrato->Valor_Adiciones + $contrato->Valor_Inicial - $contrato->Pagado, 0, '', '.'), 1, 0, 'R');
    $pdf->Cell(27, 5, utf8_decode($contrato->Estado), 1, 1, 'L');
    $cont++;
}//Fin foreach

/*
 * Creacion de las graficas
 *
 */


/*
//Pie chart
$pdf->SetFont('Helvetica', '', 12);
$pdf->Cell(0, 5, '1 - Pie chart', 0, 1);
$pdf->Ln(8);
*/
//Datos a mostrar
$pdf->SetFont('Helvetica', '', 10);
$valX = $pdf->GetX();
$valY = $pdf->GetY();
$pdf->Ln();

//Recorrido de la informacion
foreach ($this->informes_model->contar_estados() as $estado) {
    
    $pdf->Cell(30, 5, utf8_decode($estado->Estado));
    $pdf->Cell(15, 5, number_format($estado->Total, 0, '' , '.') , 0, 0, 'R');
    $pdf->Ln();
    $data[utf8_decode($estado->Estado)] = $estado->Total;
}

$valY = $pdf->GetY();

/*
$pdf->Ln();
$pdf->Cell(30, 5, 'Number of children:');
$pdf->Cell(15, 5, number_format($data['Children'], 0, '' , '.') , 0, 0, 'R');
*/

//Colores de cada sector del diagrama
$pdf->SetXY(5, $valY);
$col1=array(78,247,112);
$col2=array(245,8,8);
$col3=array(255,255,100);

$colores = array($col1,$col2);

//
$pdf->PieChart(100, 70, $data, '%l (%p)', $colores);
$pdf->SetXY($valX, $valY + 40);

/*
//Bar diagram
$pdf->SetFont('Arial', 'BIU', 12);
$pdf->Cell(0, 5, '2 - Bar diagram', 0, 1);
$pdf->Ln(8);
$valX = $pdf->GetX();
$valY = $pdf->GetY();
$pdf->BarDiagram(190, 70, $data, '%l : %v (%p)', array(255,175,100));
$pdf->SetXY($valX, $valY + 80);
*/
$nombre = str_replace(' ', '_', $nombre);
$pdf->Output($nombre, 'I');
?>
