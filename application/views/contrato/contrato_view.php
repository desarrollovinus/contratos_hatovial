<?php 
//Atributos para el fieldset
$fieldset = array('id' => 'fieldset', 'class' => 'fieldset');

//Se inicia el formulario
echo form_open('contrato/agregar_contrato');
?>

<div id="form" class="container_11">
    <div class="titulos_formularios">Registro de contrato</div>
    
    <!--Errores de validaci&oacute;n-->
    <span class="error">
        <?php 
        echo form_error('numero_contrato');
        echo form_error('contratista');
        echo form_error('objeto_contrato');
        echo form_error('localizacion_contrato');
        echo form_error('estados_contratos');
        echo form_error('valor_inicial');
        echo form_error('plazo');
        echo form_error('porcentaje_avance');
        ?>
    </span>
    
    <div id="accordion">
        <!--Seccion 1-->
        <h3><a href="#seccion1">DATOS GENERALES</a></h3>
        <div>
            <?php 
            $_contratista = array('' => null);
            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
            foreach ($contratistas as $contratista):
                $_contratista[$contratista->Pk_Id_Terceros] = $contratista->Nombre;
            endforeach; 
            ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero de contrato*', 'numero_contrato'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_contrato', 'id' => 'numero_contrato', 'style' => 'text-align: right', 'value' => set_value('numero_contrato'))) ?></td>
                    <td><?php echo form_label('Contratista*', 'contratista'); ?></td>
                    <td><?php echo form_dropdown('contratista', $_contratista, set_value('contratista')); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Objeto del contrato*', 'objeto_contrato'); ?></td>
                    <td></td>
                    <td style="font-size: 0.8em; color: #0B37B0;"><?php echo form_label('Agregar nuevo', 'nuevo_contratista') ?></td>
                    <td><?php echo form_input(array('name' => 'nuevo_contratista', 'id' => 'nuevo_contratista', 'value' => set_value('nuevo_contratista'))); ?></td>
                </tr>
                <tr>
                    <td colspan="3"><?php echo form_textarea(array('name' => 'objeto_contrato', 'id' => 'objeto_contrato', 'value' => set_value('objeto_contrato'))); ?></td>
                    <td colspan="2">
                        <table width="100%">
                            <tr>
                                <td><?php echo form_label('Localizaci&oacute;n*', 'localizacion_contrato'); ?></td>
                                <td><?php echo form_input(array('name' => 'localizacion_contrato', 'id' => 'localizacion_contrato', 'value' => set_value('localizacion_contrato'))); ?></td>
                            </tr>
                            <?php
                            $_estado_contrato = array('' => '');
                            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                            foreach($contratos_estados as $estados):
                                $_estado_contrato[$estados->Pk_Id_Estado] = $estados->Estado;
                            endforeach;

                            $_contratantes = array('' => '');
                            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                            foreach($contratantes as $contratante):
                                $_contratantes[$contratante->Pk_Id_Terceros] = $contratante->Nombre;
                            endforeach;
                            ?>
                            <tr>
                                <td><?php echo form_label('Estado*', 'estados_contrato'); ?></td>
                                <td><?php echo form_dropdown('estados_contratos', $_estado_contrato); ?></td>
                            </tr>
                            <tr>
                                <td><?php echo form_label('Contratante', 'contratante'); ?></td>
                                <td><?php echo form_dropdown('contratante', $_contratantes); ?></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
        
        <!--Seccion 2-->
        <h3><a href="#seccion2">VALORES DEL CONTRATO</a></h3>
        <div>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('Valor Inicial', 'valor_inicial'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_inicial', 'id' => 'valor_inicial', 'style' => 'text-align: right', 'value' => set_value('valor_inicial'))); ?></td>
                    <td><?php echo form_label('Fecha Inicial', 'fecha_inicial'); ?></td>
                    <td><?php echo form_input(array('name' => 'fecha_inicial', 'id' => 'fecha_inicial', 'readonly' => 'readonly', 'value' => set_value('fecha_inicial'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('¿Acta de inicio?', 'acta_inicio'); ?></td>
                    <td><?php echo form_dropdown('acta_inicio', array('' => '', 1 => 'Si', 0 => 'No')); ?></td>
                    <td><?php echo form_label('Fecha del acta', 'fecha_acta'); ?></td>
                    <td><?php echo form_input(array('name' => 'fecha_acta', 'id' => 'fecha_acta', 'readonly' => 'readonly','value' => set_value('fecha_acta'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Plazo (d&iacute;as)', 'plazo'); ?></td>
                    <td><?php echo form_input(array('name' => 'plazo', 'id' => 'plazo', 'style' => 'text-align:right', 'value' => set_value('plazo'))); ?></td>
                    <td><?php echo form_label('Porcentaje de avance', 'porcentaje_avance'); ?></td>
                    <td><?php echo form_input(array('name' => 'porcentaje_avance', 'id' => 'porcentaje_avance', 'style' => 'text-align:right', 'value' => set_value('porcentaje_avance'))); ?></td>
                </tr>
            </table>
        </div>
        
        <!--Seccion 3-->
        <h3><a href="#seccion3">DATOS DE LAS P&Oacute;LIZAS</a></h3>
        <div>
            <!--P&oacute;liza de cumplimiento-->
            <?php echo form_fieldset('P&oacute;liza de cumplimiento', $fieldset) ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_cumplimiento') ?></td>
                    <td><?php echo form_input(array('name' => 'numero_cumplimiento', 'id' => 'numero_cumplimiento', 'value' => set_value('numero_cumplimiento'))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_cumplimiento', 'id' => 'inicio_cumplimiento', 'readonly' => 'readonly','value' => set_value('inicio_cumplimiento'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_cumplimiento', 'id' => 'vencimiento_cumplimiento', 'readonly' => 'readonly','value' => set_value('vencimiento_cumplimiento'))); ?></td>
                    <td><?php echo form_label('Valor', 'valor_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_cumplimiento', 'id' => 'valor_cumplimiento', 'style' => 'text-align: right', 'value' => set_value('valor_cumplimiento'))); ?></td>
                </tr>
                <?php
                $_aseguradoras_cumplimiento = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_cumplimiento[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_cumplimiento'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_cumplimiento', $_aseguradoras_cumplimiento); ?></td>
                </tr>
            </table><br>
            
            <!--P&oacute;liza de prestaciones-->
            <?php echo form_fieldset('P&oacute;liza de prestaciones sociales', $fieldset) ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_prestaciones') ?></td>
                    <td><?php echo form_input(array('name' => 'numero_prestaciones', 'id' => 'numero_prestaciones', 'value' => set_value('numero_prestaciones'))); ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_prestaciones', 'id' => 'inicio_prestaciones', 'readonly' => 'readonly','value' => set_value('inicio_prestaciones'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_prestaciones', 'id' => 'vencimiento_prestaciones', 'readonly' => 'readonly','value' => set_value('vencimiento_prestaciones'))); ?></td>
                    <td><?php echo form_label('Valor', 'valor_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_prestaciones', 'style' => 'text-align: right', 'id' => 'valor_prestaciones', 'value' => set_value('valor_prestaciones'))); ?></td>
                </tr>
                <?php
                    $_aseguradoras_prestaciones = array('' => '');
                    //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                    foreach($aseguradoras as $aseguradora):
                        $_aseguradoras_prestaciones[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                    endforeach;
                    ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_prestaciones'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_prestaciones', $_aseguradoras_prestaciones); ?></td>
                </tr>
            </table>
            <?php echo form_fieldset_close(); ?>
            
            <!--P&oacute;liza de anticipos-->
            <?php echo form_fieldset('P&oacute;liza de Anticipos', $fieldset) ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_anticipos', 'id' => 'numero_anticipos', 'value' => set_value('numero_anticipos'))); ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_anticipos', 'id' => 'inicio_anticipos', 'readonly' => 'readonly','value' => set_value('inicio_anticipos'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_anticipos', 'id' => 'vencimiento_anticipos', 'readonly' => 'readonly','value' => set_value('vencimiento_anticipos'))); ?></td>
                    <td><?php echo form_label('Valor', 'valor_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_anticipos', 'id' => 'valor_anticipos', 'style' => 'text-align:right', 'value' => set_value('valor_anticipos'))); ?></td>
                </tr>
                <?php
                $_aseguradoras_anticipos = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_anticipos[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_prestaciones'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_anticipos', $_aseguradoras_anticipos); ?></td>
                </tr>
            </table>
            <?php echo form_fieldset_close(); ?>
            
            <!--P&oacute;liza de calidad-->
            <?php echo form_fieldset('P&oacute;liza de Calidad', $fieldset) ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_calidad', 'id' => 'numero_calidad', 'value' => set_value('numero_calidad'))); ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_calidad', 'id' => 'inicio_calidad', 'readonly' => 'readonly','value' => set_value('inicio_calidad'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_calidad', 'id' => 'vencimiento_calidad', 'readonly' => 'readonly','value' => set_value('vencimiento_calidad'))); ?></td>
                    <td><?php echo form_label('Valor', 'valor_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_calidad', 'id' => 'valor_calidad', 'style' => 'text-align: right', 'value' => set_value('valor_calidad'))); ?></td>
                </tr>
                <?php
                $_aseguradoras_calidad = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_calidad[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_calidad'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_calidad', $_aseguradoras_calidad); ?></td>
                </tr>
            </table>
            <?php echo form_fieldset_close(); ?>
            
            <!--P&oacute;liza de estabilidad-->
            <?php echo form_fieldset('P&oacute;liza de Estabilidad', $fieldset) ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_estabilidad', 'id' => 'numero_estabilidad', 'value' => set_value('numero_estabilidad'))); ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_estabilidad', 'id' => 'inicio_estabilidad', 'readonly' => 'readonly','value' => set_value('inicio_estabilidad'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_estabilidad', 'id' => 'vencimiento_estabilidad', 'readonly' => 'readonly','value' => set_value('vencimiento_estabilidad'))); ?></td>
                    <td><?php echo form_label('Valor', 'valor_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_estabilidad', 'id' => 'valor_estabilidad', 'style' => 'text-align: right', 'value' => set_value('valor_estabilidad'))); ?></td>
                </tr>
                <?php
                $_aseguradoras_estabilidad = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_estabilidad[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_estabilidad'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_estabilidad', $_aseguradoras_estabilidad); ?></td>
                </tr>
            </table>
            <?php echo form_fieldset_close(); ?>
            
            <!--P&oacute;liza de Responsabilidad Civil Contractual-->
            <?php echo form_fieldset('P&oacute;liza RC', $fieldset) ?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_rc', 'id' => 'numero_rc', 'value' => set_value('numero_rc'))); ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_rc', 'id' => 'inicio_rc', 'readonly' => 'readonly','value' => set_value('inicio_rc'))); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_rc', 'id' => 'vencimiento_rc', 'readonly' => 'readonly','value' => set_value('vencimiento_rc'))); ?></td>
                    <td><?php echo form_label('Valor', 'valor_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_rc', 'id' => 'valor_rc', 'style' => 'text-align: right', 'value' => set_value('valor_rc'))); ?></td>
                </tr>
                <?php
                $_aseguradoras_rc = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_rc[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_rc'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_rc', $_aseguradoras_rc); ?></td>
                </tr>
            </table>
        </div>
        <?php echo form_fieldset_close(); ?>
    </div><br>
    <div align="center">
        <?php
        echo form_submit(array( 'name' => 'guardar', 'id' => 'guardar', 'value' => 'Guardar'));
        echo form_input(array('type' => 'button', 'name' => 'volver', 'id' => 'volver', 'value' => 'Cancelar', 'onclick'=> "history.back()"));
        ?>
    </div>
    <?php
    //Se cierra el formulario
    echo form_close();
    ?>
</div>