<div class="titulos_formularios">Listado de contratos en liquidaci&oacute;n</div><br>
<table cellpadding="0" cellspacing="0" border="" class="display" id="example" style="font-size: 15px;">
    <thead> 
        <tr> 
            <th class="primero">Contrato</th> 
            <th>Estado</th> 
            <th style="text-align: center">Valor</th> 
            <th>Fecha Inicial</th> 
            <th class="ultimo" style="text-align: center">Opciones</th> 
        </tr> 
    </thead>
    <tbody>
        <?php foreach ($contratos as $contrato): ?>
        <tr>
            <td width="12%" style="text-align: right"><?php echo $contrato->Numero; ?></td>
            <td width="15%"><?php echo $contrato->Estado; ?></td>
            <td width="20%" style="text-align: right"><?php echo '$ '.number_format($contrato->Valor_Inicial,0,'','') ; ?></td>
            <td width="20%" style="text-align: right"><?php echo $this->auditoria_model->formato_fecha($contrato->Fecha_Inicial); ?></td>
            <td style="text-align: right">
                <?php
                echo anchor(site_url('informes/acta_liquidacion_word/'.$contrato->Pk_Id_Contrato), img(array('src' => 'img/word.png', 'title' => 'Generar acta de liquidación', 'width' => '30', 'height' => '30')));
                echo anchor_popup(site_url('informes/acta_liquidacion/'.$contrato->Pk_Id_Contrato), img(array('src' => 'img/pdf.png', 'title' => 'Generar acta de liquidación')), array('width' => '800','height' => '600','scrollbars' => 'yes','status' => 'yes','resizable' => 'yes','screenx' => '0','screeny' => '0'));
                echo anchor(site_url('informes/acta_recibo/'.$contrato->Pk_Id_Contrato), img(array('src' => 'img/word.png', 'title' => 'Generar acta de recibo', 'width' => '30', 'height' => '30')));
                ?>
            </td>
        </tr>
        <?php endforeach; ?>
    </tbody>
</table>

<script type='text/javascript'>
     $(document).ready(function(){
         /************************Scripts para las tablas************************/
        $('#example').dataTable({
            "bPaginate": true,
            "bLengthChange": true,
            "bFilter": true,
            "bSort": true,
            "bInfo": true,
            'fillSpace': true,
            "bAutoWidth": true,

            //Este script establece un orden por cierta columna
            "aaSorting": [[ 0, "asc" ]]
        });
     })
</script>