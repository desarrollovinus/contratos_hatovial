
<?php
//Se obtiene el id del contrato
$id_contrato = $this->uri->segment(3);

//Se carga el modelo que trae toda la infoormaci&oacute;n del contrato
$contratos = $this->contrato_model->listar_contratos($id_contrato);

//Se traen los pagos existentes de ese contrato
$pagos = $this->pago_model->listar_pagos($id_contrato);

/*
 * Creaci&oacute;n del objeto de la clase heredada
 * Informaci&oacute;n general del informe
 */
$pdf = new PDF();
$pdf->AliasNbPages();
$pdf->AddPage();
$pdf->SetMargins(20, 20, 20);
$pdf->SetAuthor('John Arley Cano - Hatovial S.A.S.');
$pdf->SetTitle('John Arley Cano - Hatovial S.A.S.');
$pdf->SetCreator('John Arley Cano - Hatovial S.A.S.');

foreach ($contratos as $contrato):
    $pdf->SetFont('Arial','B',12);
    $pdf->MultiCell(0,6, utf8_decode('ACTA DE LIQUIDACIÓN'),0,'C');
    $pdf->Ln(10);

    $pdf->SetFont('Arial','',12);
    $texto1 = file_get_contents('application/views/informes/plantilla1.txt');
    $texto1 = str_replace('{CONTRATO}', $contrato->Numero, $texto1);
    $texto1 = str_replace('{CONTRATISTA}', $contrato->Contratista, $texto1);
    $texto1 = str_replace('{DOCUMENTO_CONTRATISTA}', $contrato->Documento_Contratista, $texto1);
    $texto1 = str_replace('{REPRESENTANTE_LEGAL}', strtoupper($contrato->Representante_Legal), $texto1);

    /*
     * Cuerpo del mensaje
     */
    //Inicio
    $pdf->MultiCell(0,6, utf8_decode($texto1),0,'J');
    $pdf->Ln(10);

    //Objeto
    $pdf->SetFont('Arial','B',12);
    $pdf->MultiCell(0,6, utf8_decode('OBJETO DEL CONTRATO'),0,'J');
    $pdf->SetFont('Arial','',12);
    $pdf->MultiCell(0,6, utf8_decode($contrato->Objeto),0,'J');
    $pdf->Ln(10);

    $numero = new NumerosALetras();
    $pdf->SetFont('Arial','B',12);
    $pdf->MultiCell(0,6, utf8_decode('VALOR DEL CONTRATO'),0,'J');
    $pdf->SetFont('Arial','',12);
    $pdf->MultiCell(0,6, utf8_decode('$ '.number_format($contrato->Valor_Inicial,2,',','.').' ('.$numero->traducir($contrato->Valor_Inicial).')'),0,'J');
    $pdf->Ln(10);

    $pdf->SetFont('Arial','B',12);
    $pdf->MultiCell(0,6, utf8_decode('FECHA DE INICIACIÓN: '.$this->auditoria_model->formato_fecha($contrato->Fecha_Inicial)),0,'J');
    $pdf->Ln();
    $pdf->MultiCell(0,6, utf8_decode('FECHA DE VENCIMIENTO: '.$this->auditoria_model->formato_fecha($contrato->Fecha_Vencimiento)),0,'J');
    $pdf->Ln(10);

    $pdf->SetFont('Arial','B',12);
    $pdf->MultiCell(0,6, utf8_decode('VALOR TOTAL EJECUTADO POR EL CONTRATISTA'),0,'J');
    $pdf->SetFont('Arial','',12);
    $pdf->MultiCell(0,6, utf8_decode('$ '.number_format($contrato->Valor_Inicial,2,',','.').' ('.$numero->traducir($contrato->Valor_Inicial).')'),0,'J');
    $pdf->Ln(10);

    $pdf->SetFont('Arial','',12);
    $texto2 = file_get_contents('application/views/informes/plantilla2.txt');
    $texto2 = str_replace('{FECHA_VENCIMIENTO}', $this->auditoria_model->formato_fecha($contrato->Fecha_Vencimiento), $texto2);
    $texto2 = str_replace('{REPRESENTANTE_LEGAL}', strtoupper($contrato->Representante_Legal), $texto2);
    $texto2 = str_replace('{CONTRATISTA}', $contrato->Contratista, $texto2);
    $pdf->MultiCell(0,6, utf8_decode($texto2),0,'J');
    $pdf->Ln(10);

    $pdf->SetFont('Arial','',12);
    $numero = 1;
    foreach ($pagos as $pago):
        if($numero == 1){
            $pdf->SetFont('Arial','B',12);
            $pdf->MultiCell(0,6, utf8_decode('BALANCE GENERAL DEL CONTRATO'),0,'J');
            $pdf->Ln();

            $pdf->SetFont('Arial','',12);
            $pdf->MultiCell(0,6, utf8_decode('A continuación se presentará el balance financiero de la oferta'),0,'J');
            $pdf->Ln(10);
        }
        $pdf->MultiCell(0,6, 'VALOR FACTURADO EN ACTA '.$pago->Numero_Acta.' ('.$this->auditoria_model->formato_fecha($pago->Fecha).')     $ '.number_format($pago->Valor_Pago, 0, '', '.'),0,'J');
        $pdf->MultiCell(0, 6, utf8_decode('Rete garantía (5% antes de IVA)                                     $ '.number_format($pago->Valor_Retenido, 0, '', '.')));
        $pdf->MultiCell(0, 6, utf8_decode('FACTURA '.$numero.' ('.$contrato->Contratista.')'));
        $numero++;
        $pdf->Ln(10);
    endforeach;

    $pdf->SetFont('Arial','B',12);
    $numero = 1;
    foreach ($pagos as $pago):
        if($numero == 1){

            $pdf->MultiCell(0,6, utf8_decode('RELACIÓN DE RETENIDOS'),0,'J');
            $pdf->Ln();
        }
        $pdf->SetFont('Arial','',12);
        $pdf->MultiCell(0,6, utf8_decode('Acta Factura '.$numero.': Rete garantía (5% antes del IVA) $ '.number_format($pago->Valor_Retenido, 0, '', '.')), 0,'J');
        $numero++;
    endforeach;

    $pdf->Ln(10);
    $pdf->SetFont('Arial','B',12);
    $pdf->MultiCell(0,6, utf8_decode('DOCUMENTOS APORTADOS EN LA LIQUIDACIÓN'),0,'J');
    $pdf->Ln(10);

    $pdf->SetFont('Arial','',12);
    $texto3 = file_get_contents('application/views/informes/plantilla3.txt');
    $pdf->MultiCell(0,6, utf8_decode($texto3),0,'J');
    $pdf->Ln(10);

    $pdf->MultiCell(0,6, 'Para constancia se firma por quienes en ella intervinieron.',0,'J');
    $pdf->Ln(30);

    /*
     * Firmas
     */
    $pdf->SetFont('Arial','B',12);
    $pdf->Cell(0,6, utf8_decode('Ing. RICARDO LÓPEZ LOMBANA'),0,'J');

    $pdf->SetX(105);
    $pdf->Cell(0,6, utf8_decode('Ing. '.strtoupper($contrato->Representante_Legal)),0,'J');
    $pdf->Ln();

    $pdf->SetFont('Arial','',12);
    $pdf->Cell(0,6, utf8_decode('Representante Legal'),0,'J');

    $pdf->SetX(105);
    $pdf->SetFont('Arial','',12);
    $pdf->Cell(0,6, utf8_decode('Representante Legal'),0,'J');
    $pdf->Ln();

    $pdf->SetFont('Arial','',12);
    $pdf->Cell(0,6, utf8_decode('Consorcio Constructor Aburrá Norte'),0,'J');

    $pdf->SetX(105);
    $pdf->SetFont('Arial','',12);
    $pdf->Cell(0,6, utf8_decode($contrato->Contratista),0,'J');
    $pdf->Ln();

    $pdf->SetFont('Arial','',12);
    $pdf->Cell(0,6, utf8_decode('COCAN'),0,'J');

    //Se Imprime el informe en PDF
    $pdf->Output('Acta de liquidacion contrato '.$contrato->Numero, 'I');
endforeach;
?>