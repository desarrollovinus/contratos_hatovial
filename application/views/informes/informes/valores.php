<?php
//Se listan los contratos con la informacion necesaria
$contratos = $this->informes_model->listar_valores(null);

/*
 * Creaci&oacute;n del objeto de la clase heredada
 * Informaci&oacute;n general del informe
 */
$pdf = new FPDF();
$pdf->AliasNbPages();
$pdf->AddPage();
$pdf->SetAuthor('John Arley Cano - Hatovial S.A.S.');
$pdf->SetTitle('John Arley Cano - Hatovial S.A.S.');
$pdf->SetCreator('John Arley Cano - Hatovial S.A.S.');

//Cabecera
$pdf->Image('img/logo2.png', 10, 5);
$pdf->SetFont('Arial','B',14);
$pdf->Cell(190, 7, utf8_decode('Valores de los contratos'), 0, 0, 'C');
$pdf->Ln(12);

$pdf->SetFont('Arial','B',12);
$pdf->Cell(20, 7, utf8_decode('Número'), 1, 0, 'C');
$pdf->Cell(80, 7,'Contratista', 1, 0, 'C');
$pdf->Cell(30, 7,'Valor Inicial', 1, 0, 'C');
$pdf->Cell(30, 7,'Valor Pagado', 1, 0, 'C');
$pdf->Cell(30, 7,'Saldo', 1, 0, 'C');
$pdf->Ln();

//Contenido
$pdf->SetFont('Arial','',10);

//Recorrido por los contratos
foreach ($contratos as $contrato) {
	$pdf->Cell(20, 7, $contrato->Numero, 1, 0, 'R');
	$pdf->Cell(80, 7, substr(utf8_decode($contrato->Contratista), 0, 35), 1, 0, 'L');
	$pdf->Cell(30, 7, '$ '.number_format($contrato->Valor_Inicial, 0, '', '.'), 1, 0, 'R');
	// $pdf->Cell(30, 7, '$ '.number_format($contrato->Valor_Pagado, 0, '', '.'), 1, 0, 'R');
	// $pdf->Cell(30, 7, '$ '.number_format($contrato->Saldo, 0, '', '.'), 1, 0, 'R');
	$pdf->Ln();	
}//Fin foreach

//$pdf->Output('Valores_Contratos.pdf', 'I');
$pdf->Output('Valores_Contratos.pdf', 'D');