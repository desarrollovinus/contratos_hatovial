<?php
$fieldset = array('id' => 'fieldset', 'class' => 'fieldset');
//Se inicia el formulario
echo form_open('bitacora/agregar/'.$id_contrato);
echo form_hidden('id_contrato', $id_contrato);
?>

<div id="form" class="container_11">
    <div class="titulos_formularios">Agregar nueva anotaci&oacute;n</div><br>
    <!--Errores de validaci&oacute;n-->
    <span class="error">
        <?php
        echo form_error('asunto');
        ?>
    </span>
    <table>
        <tr>
            <td><?php echo form_label('Asunto*', 'asunto') ?></td>
            <td><?php echo form_input(array('name' => 'asunto', 'id' => 'asunto', 'value' => set_value('asunto'))); ?></td>
        </tr>
        <tr>
            <td><?php echo form_label('Observaci&oacute;n&nbsp;&nbsp;', 'observacion'); ?></td>
        </tr>
        <tr>
            <td colspan="2"><?php echo form_textarea(array('name' => 'observacion', 'id' => 'observacion', 'style' => /*'width:870px;*/ 'font-family: Tahoma; font-size: 14px;', 'value' => set_value('observacion'))); ?></td> 
        </tr>
        <tr>
            <td>
                <div>
                    <?php
                    echo form_submit(array('id' => 'guardar', 'name' => 'guardar'), 'Guardar');
                    echo form_input(array('type' => 'button', 'name' => 'volver', 'id' => 'volver', 'value' => 'Regresar', 'onclick'=> 'history.back()'));
                    ?>
                </div>
            </td>
        </tr>
    </table>
</div>
<?php
//Se cierra el formulario
echo form_close();
echo form_fieldset('', $fieldset);
echo form_fieldset_close();
?>

<table cellpadding="0" cellspacing="0" border="" class="display" id="example" style="font-size: 15px;">
    <thead>
        <tr>
            <th class="primero" width="30%" style="text-align: center">Fecha</th>
            <th style="text-align: center" width="20%">Usuario</th>
            <th style="text-align: center" width="15%">Asunto</th>
            <th class="ultimo" style="text-align: center" width="35%">Observaci&oacute;n</th>
        </tr>
    </thead>
    <tbody>
    <?php foreach($bitacoras as $bitacora): ?>
        <tr>
            <td><?php echo $this->auditoria_model->formato_fecha($bitacora->Fecha); ?></td>
            <td><?php echo $bitacora->Nombres.' '.$bitacora->Apellidos ?></td>
            <td><?php echo $bitacora->Asunto; ?></td>
            <td><?php echo $bitacora->Observacion; ?></td>
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