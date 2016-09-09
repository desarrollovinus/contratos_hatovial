<?php
// $numero = count($contratistas);
// echo "para {$nombre} hay {$numero} contratos";

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

//Cabecera
$pdf->Image('img/logo2.png', 5, 5);
$pdf->SetFont('Arial','B',14);
$pdf->Cell(340, 7, utf8_decode(substr("CONTRATOS ASOCIADOS A " . $nombre, 0, 75)), 0, 1, 'C');
$pdf->Ln(10);














$nombre = str_replace(' ', '_', $nombre);
$pdf->Output('Contratos_' . $nombre, 'I');
?>