<?php
$fieldset = array('id' => 'fieldset', 'class' => 'fieldset');
echo form_open_multipart('pago/agregar/'.$id_contrato);
echo form_hidden('id_contrato', $id_contrato);
?>

<script language="javascript" src="jquery-1.2.6.min.js"></script>

<?php if($this->session->userdata('Tipo') == TRUE){ ?>
<div id="form" class="container_12">
    <!--Errores de validaci&oacute;n-->
    <span class="error">
        <?php
        echo form_error('valor');
        echo form_error('retenido');
        ?>
    </span>
    <table width="100%">
        <tr>
            <td><?php echo form_label('Fecha', 'fecha_pago') ?></td>
            <td><?php echo form_input(array('name' => 'fecha_pago', 'id' => 'fecha_pago', 'readonly' => 'readonly'), set_value('fecha_pago')) ?></td>
            <td rowspan="5">
                <table width="80%">
                    <?php
                    foreach ($estado_pagos as $estado):
                    echo form_hidden('numero_contrato', $estado->Numero);
                    ?>
                    <tr>
                        <td colspan="2" style="text-align: center"><b><u>Estado de pago del contrato <?php echo $estado->Numero; ?></u></b></th>
                    </tr>
                    <tr>
                        <td><b>Valor (incluídas adiciones): </b></td>
                        <td style="text-align: right"><?php echo '$ '.number_format($estado->Valor_Inicial + $estado->Valor_Adiciones, 0, '', '.'); ?></td>
                    </tr>
                    <tr>
                        <td><b>Total pagado: </b></td>
                        <td style="text-align: right"><?php echo '$ '.number_format($estado->Pagado, 0, '', '.'); ?></td>
                    </tr>
                    <tr>
                        <td><b>Saldo: </b></td>
                        <td style="text-align: right"><?php echo '$ '.number_format($estado->Valor_Inicial + $estado->Valor_Adiciones - $estado->Pagado, 0, '', '.'); ?></td>
                    </tr>
                    <tr>
                        <td><b>Retenido: </b></td>
                        <td style="text-align: right"><?php echo '$ '.number_format($estado->Valor_Retenido, 0, '', '.'); ?></td>
                    </tr>
                    <tr>
                        <td><b>Porcentaje:  </b></td>
                        <td style="text-align: right"><?php echo number_format($estado->Porcentaje, 2, '.', '').'%'; ?></td>
                    </tr>
                    <?php endforeach; ?>
                </table>
            </td>
        </tr>
        <tr>
            <td><?php echo form_label('N&uacute;mero de acta', 'numero_acta'); ?></td>
            <td><?php echo form_input(array('name' => 'numero_acta', 'id' => 'numero_acta', 'style' => 'text-align:right'), set_value('numero_acta')); ?></td>
        </tr>
        <tr>
            <td><?php echo form_label('Valor*', 'valor'); ?></td>
            <td><?php echo form_input(array('name' => 'valor', 'id' => 'valor', 'style' => 'text-align:right'), set_value('valor')); ?></td>
        </tr>
        <tr>
            <td><?php echo form_label('Valor retenido', 'retenido'); ?></td>
            <td><?php echo form_input(array('name' => 'retenido', 'id' => 'retenido', 'style' => 'text-align:right'), set_value('retenido')); ?></td>
        </tr>
        <tr>
            <td colspan="2">
                <div align="left">
                    <?php
                    echo form_submit(array('id' => 'guardar', 'name' => 'guardar'), 'Generar pago');
                    echo form_input(array('type' => 'button', 'name' => 'volver', 'id' => 'volver', 'value' => 'Regresar', 'onclick'=> 'history.back()'));
                    ?>
                </div>
            </td>
        </tr>
    </table>
</div>
<?php
echo form_close();
echo form_fieldset('', $fieldset);
echo form_fieldset_close();
}
?>

<table cellpadding="0" cellspacing="0" border="" class="display" id="example" style="font-size: 15px;">
    <thead>
        <tr>
            <th width="5%" class="primero">Pago</th>
            <th width="30%">Fecha</th>
            <th width="20%">N&uacute;mero de acta</th>
            <th width="20%">Valor del pago</th>
            <th width="20%">Valor retenido</th>
            <th width="5%" class="ultimo">Opciones</th>
        </tr>
    </thead>
    <tbody>
        <?php
        $numero_pago = 1;
        foreach ($pagos as $pago):
        ?>
        <tr>
            <td style="text-align: right"><?php echo $numero_pago; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($pago->Fecha); ?></td>
            <td style="text-align: left"><?php echo $pago->Numero_Acta; ?></td>
            <td style="text-align: right"><?php echo '$ '.number_format($pago->Valor_Pago, 0, '', '.'); ?></td>
            <td style="text-align: right"><?php echo '$ '.number_format($pago->Valor_Retenido, 0, '', '.'); ?></td>
            <?php if($this->session->userdata('Tipo') == TRUE){ ?>
            <td style="text-align: right"><?php echo anchor(site_url("pago/eliminar/".$pago->Pk_Id_Contratos_Pagos."/".$pago->Pk_Id_Contrato), img(array('src' => 'img/borrar.png', 'title' => 'Eliminar pago', 'width' => '25', 'height' => '25'))); ?></td>
            <?php }else{ ?>
            <td></td>
            <?php } ?>
        </tr>
        <?php
        $numero_pago++;
        endforeach;
        ?>
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