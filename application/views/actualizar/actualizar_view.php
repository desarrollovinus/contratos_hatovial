<?php
foreach ($contratos as $contrato):
    
//Seccion 1
$fieldset = array('id' => 'fieldset', 'class' => 'fieldset');
 
echo form_open('actualizar/contrato/'.$contrato->Pk_Id_Contrato);
?>

<div class="container_12" id="form">
    <div class="titulos_formularios">Modificaci&oacute;n de contrato</div>
    
    <!--Errores de validaci&oacute;n-->
    <span class="error">
        <?php
        echo form_error('contratista');
        echo form_error('objeto_contrato');
        echo form_error('localizacion_contrato');
        echo form_error('estados_contratos');
        echo form_error('valor_inicial');
        echo form_error('plazo');
        echo form_error('valor_cumplimiento');
        echo form_error('valor_prestaciones');
        echo form_error('valor_anticipos');
        echo form_error('valor_calidad');
        echo form_error('valor_estabilidad');
        echo form_error('plazo_adicion');
        echo form_error('valor_adicion');
        ?>
    </span>
    
    <div id="accordion">
        <!--Seccion 1-->
        <h3><a href="#seccion1">DATOS GENERALES</a></h3>
        <div>
            <table width="100%">
                <?php
                    //Con un campo oculto obtenemos el id del contrato
                    echo form_hidden('id_contrato', $contrato->Pk_Id_Contrato);
                    $_contratista = array('' => null);
                    
                    //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                    foreach ($contratistas as $contratista):
                        $_contratista[$contratista->Pk_Id_Terceros] = $contratista->Nombre;
                    endforeach; 
                ?>
                <tr>
                    <td><?php echo form_label('N&uacute;mero de contrato', 'numero_contrato'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_contrato', 'id' => 'numero_contrato','style' => 'text-align:right', 'readonly' => 'readonly'), $contrato->Numero) ?></td>
                    <td><?php echo form_label('Contratista*', 'contratista'); ?></td>
                    <td><?php echo form_dropdown('contratista', $_contratista,  set_value('contratista', $contrato->Pk_Id_Terceros)); ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Objeto del contrato*', 'objeto_contrato'); ?></td>
                    <td></td>
                    <td style="font-size: 0.8em; color: #0B37B0;"><?php echo form_label('Agregar nuevo', 'nuevo_contratista') ?></td>
                    <td><?php echo form_input(array('name' => 'nuevo_contratista', 'id' => 'nuevo_contratista', 'value' => set_value('nuevo_contratista'))); ?></td>
                </tr>
                <tr>
                    <td colspan="2"><?php echo form_textarea(array('name' => 'objeto_contrato', 'id' => 'objeto_contrato', 'value' => set_value('objeto_contrato', $contrato->Objeto))) ?></td>
                    <td colspan="2">
                        <table width="100%">
                            <tr>
                                <td><?php echo form_label('Localizaci&oacute;n*', 'localizacion_contrato'); ?></td>
                                <td><?php echo form_input(array('name' => 'localizacion_contrato', 'id' => 'localizacion_contrato', 'value' => set_value('localizacion_contrato', $contrato->Lugar))) ?></td>
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
                                <td><?php echo form_label('Estado*', 'estados_contratos'); ?></td>
                                <td><?php echo form_dropdown('estados_contratos', $_estado_contrato, set_value('estados_contratos', $contrato->Pk_Id_Estado)) ?></td>
                            </tr>
                            <tr>
                                <td><?php echo form_label('Contratante', 'contratante'); ?></td>
                                <td><?php echo form_dropdown('contratante', $_contratantes, set_value('contratante', $contrato->Pk_Id_Tercero_Contratante)); ?></td>
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
                    <td><?php echo form_input(array('name' => 'valor_inicial', 'id' => 'valor_inicial', 'style' => 'text-align:right', 'value' => set_value('valor_inicial', number_format($contrato->Valor_Inicial,0,'','')))) ?></td>
                    <td><?php echo form_label('Fecha Inicial', 'fecha_inicial'); ?></td>
                    <td><?php echo form_input(array('name' => 'fecha_inicial', 'id' => 'fecha_inicial', 'readonly' => 'readonly', 'value' => set_value('fecha_inicial', $contrato->Fecha_Inicial))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('¿Acta de inicio?', 'acta_inicio'); ?></td>
                    <?php
                    $acta_inicio = array('' => '', 1 => 'Si', 0 => 'No');
                    ?>
                    <td><?php echo form_dropdown('acta_inicio', $acta_inicio, set_value('acta_inicio',$contrato->Acta_Inicio)); ?></td>
                    <td><?php echo form_label('Fecha del acta', 'fecha_acta'); ?></td>
                    <td><?php echo form_input(array('name' => 'fecha_acta', 'id' => 'fecha_acta', 'value' => set_value('fecha_acta', $contrato->Fecha_Acta_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Plazo (d&iacute;as)', 'plazo'); ?></td>
                    <td><?php echo form_input(array('name' => 'plazo', 'id' => 'plazo','style' => 'text-align:right', 'value' => set_value('plazo', $contrato->Plazo_Inicial))); ?></td>
                    <td><?php echo form_label('Porcentaje de avance', 'porcentaje_avance'); ?></td>
                    <td><?php echo form_input(array('name' => 'porcentaje_avance', 'id' => 'porcentaje_avance', 'style' => 'text-align:right', 'value' => set_value('porcentaje_avance', $contrato->Porcentaje_Avance))); ?></td>
                </tr>
            </table>
        </div>
        
        <!--Seccion 3-->
        <h3><a href="#seccion3">DATOS DE LAS P&Oacute;LIZAS</a></h3>
        <div>
            <!--P&oacute;liza de cumplimiento-->
            <?php echo form_fieldset('P&oacute;liza de cumplimiento', $fieldset) ?>
            <?php foreach ($poliza_cumplimiento as $cumplimiento):?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_cumplimiento', 'id' => 'numero_cumplimiento', 'value' => set_value('numero_cumplimiento', $cumplimiento->Numero))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_cumplimiento', 'id' => 'inicio_cumplimiento', 'readonly' => 'readonly', 'value' => set_value('inicio_cumplimiento', $cumplimiento->Fecha_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_cumplimiento', 'id' => 'vencimiento_cumplimiento', 'readonly' => 'readonly', 'value' => set_value('vencimiento_cumplimiento', $cumplimiento->Fecha_Final))) ?></td>
                    <td><?php echo form_label('Valor', 'valor_cumplimiento'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_cumplimiento', 'id' => 'valor_cumplimiento', 'style' => 'text-align:right', 'value' => set_value('valor_cumplimiento', number_format($cumplimiento->Valor, 0, '', '') ))) ?></td>
                </tr>
                <?php
                    //Con un campo oculto obtenemos el id del contrato
                    echo form_hidden('id_cumplimiento', $cumplimiento->Pk_Id_Contratos_Poliza);
                    $_aseguradoras_cumplimiento = array('' => '');
                    //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                    foreach($aseguradoras as $aseguradora):
                        $_aseguradoras_cumplimiento[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                    endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_cumplimiento'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_cumplimiento', $_aseguradoras_cumplimiento, set_value('aseguradora_cumplimiento',$cumplimiento->Pk_Id_Terceros)) ?></td>
                </tr>
            </table><br>
            <?php endforeach; ?>
            
            <!--P&oacute;liza de prestaciones-->
            <?php echo form_fieldset('P&oacute;liza de prestaciones sociales', $fieldset) ?>
            <?php foreach ($poliza_prestaciones as $prestaciones):?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_prestaciones', 'id' => 'numero_prestaciones', 'value' => set_value('numero_prestaciones', $prestaciones->Numero))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_prestaciones', 'id' => 'inicio_prestaciones', 'readonly' => 'readonly', 'value' => set_value('inicio_prestaciones', $prestaciones->Fecha_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_prestaciones', 'id' => 'vencimiento_prestaciones', 'readonly' => 'readonly', 'value' => set_value('vencimiento_prestaciones', $prestaciones->Fecha_Final))) ?></td>
                    <td><?php echo form_label('Valor', 'valor_prestaciones'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_prestaciones', 'id' => 'valor_prestaciones', 'style' => 'text-align:right', 'value' => set_value('valor_prestaciones', number_format($prestaciones->Valor, 0, '', '')))) ?></td>
                </tr>
                <?php
                //Con un campo oculto obtenemos el id del contrato
                echo form_hidden('id_prestaciones', $prestaciones->Pk_Id_Contratos_Poliza);
                $_aseguradoras_prestaciones = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_prestaciones[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_prestaciones'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_prestaciones', $_aseguradoras_prestaciones, set_value('aseguradora_prestaciones', $prestaciones->Pk_Id_Terceros)) ?></td>
                </tr>
            </table><br>
            <?php endforeach; ?>
            
            <!--P&oacute;liza de anticipos-->
            <?php echo form_fieldset('P&oacute;liza de anticipos', $fieldset) ?>
            <?php echo form_fieldset_close(); ?>
            
            <?php foreach ($poliza_anticipos as $anticipos):?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_anticipos', 'id' => 'numero_anticipos', 'value' => set_value('numero_anticipos', $anticipos->Numero))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_anticipos', 'id' => 'inicio_anticipos', 'readonly' => 'readonly', 'value' => set_value('inicio_anticipos', $anticipos->Fecha_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_anticipos', 'id' => 'vencimiento_anticipos', 'readonly' => 'readonly', 'value' => set_value('vencimiento_anticipos', $anticipos->Fecha_Final))) ?></td>
                    <td><?php echo form_label('Valor', 'valor_anticipos'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_anticipos', 'id' => 'valor_anticipos', 'style' => 'text-align:right', 'value' => set_value('valor_anticipos', number_format($anticipos->Valor, 0, '', '')))) ?></td>
                </tr>
                <?php
                //Con un campo oculto obtenemos el id del contrato
                echo form_hidden('id_anticipos', $anticipos->Pk_Id_Contratos_Poliza);
                $_aseguradoras_anticipos = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_anticipos[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_anticipos'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_anticipos', $_aseguradoras_anticipos, set_value('aseguradora_anticipos', $anticipos->Pk_Id_Terceros)) ?></td>
                </tr>
            </table>
            <?php endforeach; ?>
            
            <!--P&oacute;liza de calidad-->
            <?php echo form_fieldset('P&oacute;liza de calidad', $fieldset) ?>
            <?php echo form_fieldset_close(); ?>
            
            <?php foreach ($poliza_calidad as $calidad):?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_calidad', 'id' => 'numero_calidad', 'value' => set_value('numero_calidad', $calidad->Numero))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_calidad', 'id' => 'inicio_calidad', 'readonly' => 'readonly', 'value' => set_value('inicio_calidad', $calidad->Fecha_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_calidad', 'id' => 'vencimiento_calidad', 'readonly' => 'readonly', 'value' => set_value('vencimiento_calidad', $calidad->Fecha_Final))) ?></td>
                    <td><?php echo form_label('Valor', 'valor_calidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_calidad', 'id' => 'valor_calidad', 'style' => 'text-align:right', 'value' => set_value('valor_calidad', number_format($calidad->Valor, 0, '', '')))) ?></td>
                </tr>
                <?php
                //Con un campo oculto obtenemos el id del contrato
                echo form_hidden('id_calidad', $calidad->Pk_Id_Contratos_Poliza);
                $_aseguradoras_calidad = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_calidad[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_calidad'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_calidad', $_aseguradoras_calidad, set_value('aseguradora_calidad',$calidad->Pk_Id_Terceros)) ?></td>
                </tr>
            </table>
            <?php endforeach; ?>
            
            <!--P&oacute;liza de Estabilidad-->
            <?php echo form_fieldset('P&oacute;liza de estabilidad', $fieldset) ?>
            <?php echo form_fieldset_close(); ?>
            
            <?php foreach ($poliza_estabilidad as $estabilidad):?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_estabilidad', 'id' => 'numero_estabilidad', 'value' => set_value('numero_estabilidad', $estabilidad->Numero))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_estabilidad', 'id' => 'inicio_estabilidad', 'readonly' => 'readonly', 'value' => set_value('inicio_estabilidad', $estabilidad->Fecha_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_estabilidad', 'id' => 'vencimiento_estabilidad', 'readonly' => 'readonly', 'value' => set_value('vencimiento_estabilidad', $estabilidad->Fecha_Final))) ?></td>
                    <td><?php echo form_label('Valor', 'valor_estabilidad'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_estabilidad', 'id' => 'valor_estabilidad', 'style' => 'text-align:right', 'value' => set_value('valor_estabilidad', number_format($estabilidad->Valor, 0, '', '')))) ?></td>
                </tr>
                <?php
                //Con un campo oculto obtenemos el id del contrato
                echo form_hidden('id_estabilidad', $estabilidad->Pk_Id_Contratos_Poliza);
                $_aseguradoras_estabilidad = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_estabilidad[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_estabilidad'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_estabilidad', $_aseguradoras_estabilidad, set_value('aseguradora_estabilidad', $estabilidad->Pk_Id_Terceros)) ?></td>
                </tr>
            </table>
            <?php endforeach; ?>
            
            <!--P&oacute;liza de Responsabilidad Civil Contractual-->
            <?php echo form_fieldset('P&oacute;liza RC', $fieldset) ?>
            <?php echo form_fieldset_close(); ?>
            
            <?php foreach ($poliza_rc as $rc):?>
            <table width="100%">
                <tr>
                    <td><?php echo form_label('N&uacute;mero', 'numero_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'numero_rc', 'id' => 'numero_rc', 'value' => set_value('numero_rc', $rc->Numero))) ?></td>
                    <td><?php echo form_label('Fecha de inicio', 'inicio_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'inicio_rc', 'id' => 'inicio_rc', 'readonly' => 'readonly', 'value' => set_value('inicio_rc', $rc->Fecha_Inicio))) ?></td>
                </tr>
                <tr>
                    <td><?php echo form_label('Fecha de vencimiento', 'vencimiento_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'vencimiento_rc', 'id' => 'vencimiento_rc', 'readonly' => 'readonly', 'value' => set_value('vencimiento_rc', $rc->Fecha_Final))) ?></td>
                    <td><?php echo form_label('Valor', 'valor_rc'); ?></td>
                    <td><?php echo form_input(array('name' => 'valor_rc', 'id' => 'valor_rc', 'style' => 'text-align:right', 'value' => set_value('valor_rc', number_format($rc->Valor, 0, '', '')))) ?></td>
                </tr>
                <?php
                //Con un campo oculto obtenemos el id del contrato
                echo form_hidden('id_rc', $rc->Pk_Id_Contratos_Poliza);
                $_aseguradoras_rc = array('' => '');
                //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                foreach($aseguradoras as $aseguradora):
                    $_aseguradoras_rc[$aseguradora->Pk_Id_Terceros] = $aseguradora->Nombre;
                endforeach;
                ?>
                <tr>
                    <td><?php echo form_label('Aseguradora', 'aseguradora_rc'); ?></td>
                    <td><?php echo form_dropdown('aseguradora_rc', $_aseguradoras_rc, set_value('aseguradora_rc', $rc->Pk_Id_Terceros)) ?></td>
                </tr>
            </table>
            <?php endforeach; ?>
        </div>

        <!--Seccion 4-->
        <h3><a href="#seccion1">ADICIONES AL CONTRATO</a></h3>
        <div>
            <table width="70%">
                <thead>
                    <th class="right">Otros&iacute;</th>
                    <th class="right"><?php echo form_label('Valor', 'valor_adicion'); ?></th>
                    <th class="right"><?php echo form_label('Plazo (d&iacute;as)', 'plazo_adicion'); ?></th>
                </thead>
                <?php
                $plazo_adiciones = 0;
                $num = 1;
                foreach ($adiciones as $adicion):
                $plazo_adiciones = $plazo_adiciones + $adicion->Plazo;
                ?>
                <tr>
                    <td class="right"><?php echo $num; ?></td>
                    <td class="right"><?php echo '$ '.number_format($adicion->Valor, 0, '', '.'); ?></td>
                    <td class="right"><?php echo $adicion->Plazo; ?></td>
                </tr>
                <?php
                $num++;
                endforeach;

                echo form_hidden('plazo_adiciones', $plazo_adiciones);
                ?>
                <tr>
                    <td class="right"><?php echo $num; ?></td>
                    <td class="right"><?php echo form_input(array('name' => 'valor_adicion', 'id' => 'valor_adicion', 'style' => 'text-align:right', 'value' => set_value('valor_adicion'))); ?>
                    </td>
                    <td class="right"><?php echo form_input(array('name' => 'plazo_adicion', 'id' => 'plazo_adicion', 'style' => 'text-align:right', 'value' => set_value('plazo_adicion'))); ?>
                    </td>
                </tr>
            </table>
        </div>
    </div><br>

    <div align="center">
        <?php
        echo form_submit($guardar = array('id' => 'guardar', 'name' => 'guardar'), 'Guardar');
        echo form_input($volver = array('type' => 'button', 'name' => 'volver', 'id' => 'volver', 'value' => 'Cancelar', 'onclick'=> 'history.back()'))
        ?>
    </div>
</div>
<?php endforeach; ?>
<?php echo form_close(); ?>
