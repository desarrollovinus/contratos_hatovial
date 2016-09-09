<?php
$fieldset = array('id' => 'fieldset', 'class' => 'fieldset');
echo form_open('tercero/actualizar_tercero');
echo form_hidden('id_tercero', $id_tercero);

foreach($terceros as $tercero):
?>
<div id="form" class="container_12">
    <div class="grid_11">
        <div class="titulos_formularios">Modificar datos de <?php echo $tercero->Nombre; ?></div><br>
        <!--Errores de validaci&oacute;n-->
        <span class="error">
            <?php
            echo form_error('tipo');
            echo form_error('nombre');
            echo form_error('telefono');
            ?>
        </span>
        <table width="100%">
            <tr>
                <?php 
                $_tercero = array('' => null);
                //Recorrido para traer todos los terceros de la base de datos y agregarlos en el dropdown
                foreach ($tipo_terceros as $tipo_tercero):
                    $_tercero[$tipo_tercero->Pk_Id_Terceros_Tipo] = $tipo_tercero->Nombre;
                endforeach;
                    
                ?>
                <td><?php echo form_label('Tipo', 'tipo'); ?></td>
                <td><?php echo form_dropdown('tipo', $_tercero, set_value('tipo', $tercero->Fk_Id_Terceros_Tipo)); ?></td>
                <td><?php echo form_label('Nombre*', 'nombre') ?></td>
                <td><?php echo form_input(array('id' => 'nombre', 'name' => 'nombre', 'value' => set_value('nombre', $tercero->Nombre))) ?></td>
                <td><?php echo form_label('Documento', 'documento') ?></td>
                <td><?php echo form_input(array('id' => 'documento', 'name' => 'documento', 'value' => set_value('documento', $tercero->Documento))) ?></td>
            </tr>
            <tr>
                <td><?php echo form_label('Tel&eacute;fono', 'telefono') ?></td>
                <td><?php echo form_input(array('id' => 'telefono', 'name' => 'telefono', 'style' => 'text-align: right', 'value' => set_value('telefono', $tercero->Telefono))) ?></td>
                <td><?php echo form_label('Direcci&oacute;n', 'direccion') ?></td>
                <td><?php echo form_input(array('id' => 'direccion', 'name' => 'direccion', 'value' => set_value('direccion', $tercero->Direccion))) ?></td>
                <td><?php echo form_label('Representante legal', 'representante_legal'); ?></td>
                <td><?php echo form_input(array('id' => 'representante_legal', 'name' => 'representante_legal', 'value' => set_value('representante_legal', $tercero->Representante_Legal))); ?></td>
            </tr>
        </table><br/>
        <div align="center">
            <?php echo form_submit(array('id' => 'agregar', 'value' => 'Guardar')) ?>
            <?php echo form_input(array('id' => 'regresar', 'type' => 'button', 'value' => 'Regresar', 'onClick' => 'history.back()' )) ?>
        </div>
        <?php echo form_close(); ?>
    </div>
</div>
<?php endforeach; ?>