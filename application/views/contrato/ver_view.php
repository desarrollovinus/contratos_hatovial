<?php
//Atributos de los botones de la vista
$volver = array('type' => 'button', 'name' => 'volver', 'id' => 'volver', 'value' => 'Atrás', 'onclick'=> 'history.back()');
$modificar = array('src' => 'img/modificar.png', 'title' => 'Modificar contrato', 'width' => '40', 'height' => '40');
$pago = array('src' => 'img/pago.png', 'title' => 'Realizar pago', 'width' => '40', 'height' => '40');
$bitacora = array('src' => 'img/bitacora.png', 'title' => 'Agregar bit&aacute;cora', 'width' => '40', 'height' => '40');
$acta_inicio = array('src' => 'img/word.png', 'title' => 'Generar acta de inicio', 'width' => '40', 'height' => '40');

/*
 * Se hace el recorrido para traer el resultado, que realmente es solo una fila:
 * La correspondiente al n&uacute;mero de contrato seleccionado
 */
foreach ($contratos as $contrato):

    $valor_total = $contrato->Valor_Inicial + $contrato->Valor_Adiciones;
    $plazo_total = $contrato->Plazo_Inicial + $contrato->Plazo_Adiciones;
?>
    <center>
        <b>Contrato N&uacute;mero <?php echo $contrato->Numero; ?></b>
    </center><br>
    <div class="ver">
        <!--Informacion del contratista-->
        <div class="info_contrato">
            <h3><?php echo $contrato->Contratista; ?></h3>
            <div>
                <h4>Documento: </h4>
                <p><?php echo $contrato->Documento_Contratista; ?></p>
            </div>
            <div>
                <h4>Direcci&oacute;n: </h4>
                <p><?php echo $contrato->Direccion_Contratista; ?></p>
            </div>
            <div>
                <h4>Tel&eacute;fono: </h4>
                <p><?php echo $contrato->Telefono_Contratista; ?></p>
            </div>
            <div>
                <h4>Representante legal: </h4>
                <p><?php echo $contrato->Representante_Legal; ?></p>
            </div>
        </div>

        <!--Informaci&oacute;n general-->
        <div class="info_contrato">
            <h3>Informaci&oacute;n general del contrato</h3>
            <div>
                <h4>Objeto:</h4>
                <p><?php echo $contrato->Objeto; ?></p>
            </div>
            <div>
                <h4>Localizaci&oacute;n del objeto:</h4>
                <p><?php echo $contrato->Lugar; ?></p>
            </div>
            <div>
                <h4>Estado: </h4>
                <p><?php echo $contrato->Estado; ?></p>
            </div>
            <div>
                <h4>Contratante: </h4>
                <p><?php echo $contrato->Contratante; ?></p>
            </div>
        </div>

        <!--Valores del contrato-->
        <div class="info_contrato">
            <h3>Valores</h3>  
            <div>
                <h4>Valor Inicial:</h4>
                <p><?php echo '$ '.number_format($contrato->Valor_Inicial ,0,'','.'); ?></p>
            </div>
            <div>
                <h4>Fecha Inicial:</h4>
                <p><?php echo $this->auditoria_model->formato_fecha($contrato->Fecha_Inicial); ?></p>
            </div>
            <div>
                <h4>Valor Total:</h4>
                <p><?php echo '$ '.number_format($valor_total, 0, '', '.'); ?></p>
            </div>
            <div>
                <h4>Fecha de Vencimiento:</h4>
                <p><?php echo $this->auditoria_model->formato_fecha($contrato->Fecha_Vencimiento); ?></p>
            </div>
            <div>
                <h4>¿Acta de inicio?:</h4>
                <p><?php if( $contrato->Acta_Inicio == 1){echo 'S&iacute;';}else{echo 'No';} ?></p>
            </div>
            <div>
                <h4>Plazo:</h4>
                <p><?php echo number_format($plazo_total, 0, '', '.'); ?></p>
            </div>
            <div>
                <h4>Fecha del acta:</h4>
                <p><?php echo $this->auditoria_model->formato_fecha($contrato->Fecha_Acta_Inicio); ?></p>
            </div>
            <div>
                <h4>Porcentaje de avance:</h4>
                <p><?php if($contrato->Porcentaje_Avance){echo $contrato->Porcentaje_Avance.' %';} ?></p>
            </div>
        </div>

        <!-- Adiciones -->
        <?php if($contrato->Valor_Inicial != $valor_total || $contrato->Plazo_Inicial != $plazo_total){ ?>
            <div class="info_contrato adiciones">
                <h3>Adiciones al contrato</h3><br>
                <table width="50%">
                    <thead>
                        <th>Otros&iacute;</th>
                        <th>Valor</th>
                        <th>Plazo (d&iacute;as)</th>
                    </thead>
                    <?php
                    $num = 1;
                    foreach ($adiciones as $adicion):
                    ?>
                    <tr>
                        <td width="5%" class="right"><p><?php echo $num; ?></p></td>
                        <td class="right"><p><?php echo '$ '.number_format($adicion->Valor, 0, '', '.'); ?></p></td>
                        <td class="right"><p><?php echo $adicion->Plazo; ?></p></td>
                    </tr>
                    <?php
                    $num++;
                    endforeach;
                    ?>
                </table><br>
            </div>
        <?php } ?>
    </div>
<?php endforeach; ?>
<table id="polizas">
    <thead>
        <th>P&oacute;liza</th>
        <th>N&uacute;mero</th>
        <th>Inicio</th>
        <th>Fin</th>
        <th>Valor</th>
        <th>Aseguradora</th>
        <th class="vigencia">Vigencia (d&iacute;as)</th>
    </thead>
    <tbody>
        <?php foreach ($poliza_cumplimiento as $cumplimiento):?>
        <tr>
            <td>Cumplimiento</td>
            <td><?php echo $cumplimiento->Numero; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($cumplimiento->Fecha_Inicio); ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($cumplimiento->Fecha_Final); ?></td>
            <td class="numero"><?php echo '$ '.number_format($cumplimiento->Valor,0,'','.'); ?></td>
            <td><?php echo $cumplimiento->Nombre; ?></td>
            <td class="numero"><?php echo number_format($cumplimiento->Vigencia,0,'','.'); ?></td>
        </tr>
        <?php
        endforeach;
        foreach ($poliza_prestaciones as $prestaciones):
        ?>
        <tr>
            <td>Prestaciones</td>
            <td><?php echo $prestaciones->Numero; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($prestaciones->Fecha_Inicio); ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($prestaciones->Fecha_Final); ?></td>
            <td class="numero"><?php echo '$ '.number_format($prestaciones->Valor,0,'','.'); ?></td>
            <td><?php echo $prestaciones->Nombre; ?></td>
            <td class="numero"><?php echo number_format($prestaciones->Vigencia,0,'','.'); ?></td>
        </tr>
        <?php
        endforeach;
        foreach ($poliza_anticipos as $anticipos):
        ?>
        <tr>
            <td>Anticipos</td>
            <td><?php echo $anticipos->Numero; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($anticipos->Fecha_Inicio); ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($anticipos->Fecha_Final); ?></td>    
            <td class="numero"><?php echo '$ '.number_format($anticipos->Valor,0,'','.'); ?></td>
            <td><?php echo $anticipos->Nombre; ?></td>
            <td class="numero"><?php echo number_format($anticipos->Vigencia,0,'','.'); ?></td>
        </tr>
        <?php
        endforeach;
        foreach ($poliza_calidad as $calidad):?>
        <tr>
            <td>Calidad</td>
            <td><?php echo $calidad->Numero; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($calidad->Fecha_Inicio); ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($calidad->Fecha_Final); ?></td>
            <td class="numero"><?php echo '$ '.number_format($calidad->Valor,0,'','.'); ?></td>
            <td><?php echo $calidad->Nombre; ?></td>
            <td class="numero"><?php echo number_format($calidad->Vigencia,0,'','.'); ?></td>
        </tr>
        <?php
        endforeach;
        foreach ($poliza_estabilidad as $estabilidad):
        ?>
        <tr>
            <td>Estabilidad</td>
            <td><?php echo $estabilidad->Numero; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($estabilidad->Fecha_Inicio); ?></td>
            <td><?php echo$this->auditoria_model->formato_fecha($estabilidad->Fecha_Final); ?></td>
            <td class="numero"><?php echo '$ '.number_format($estabilidad->Valor,0,'','.'); ?></td>
            <td><?php echo $estabilidad->Nombre; ?></td>
            <td class="numero"><?php echo number_format($estabilidad->Vigencia,0,'','.'); ?></td>
        </tr>
        <?php
        endforeach;
        foreach ($poliza_rc as $rc):
        ?>
        <tr>
            <td>Responsabilidad Civil</td>
            <td><?php echo $rc->Numero; ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($rc->Fecha_Inicio); ?></td>
            <td><?php echo $this->auditoria_model->formato_fecha($rc->Fecha_Final); ?></td>
            <td class="numero"><?php echo '$ '.number_format($rc->Valor,0,'','.'); ?></td>
            <td><?php echo $rc->Nombre; ?></td>
            <td class="numero"><?php echo number_format($rc->Vigencia,0,'','.'); ?></td>
        </tr>
        <?php endforeach; ?>
    </tbody>
</table>
<?php 
echo form_fieldset('', 'class="fieldset" style="text-align: center"');
echo form_fieldset_close(); 

if($this->session->userdata('Tipo') == true){
?>
    <table width="100%">
        <tr>
            <td><?php echo form_input($volver); ?></td>
            <td>
                <?php
                echo anchor(site_url('actualizar/index/'.$contrato->Pk_Id_Contrato), img($modificar));
                echo anchor(site_url('pago/index/'.$contrato->Pk_Id_Contrato), img($pago));
                echo anchor(site_url('bitacora/index/'.$contrato->Pk_Id_Contrato), img($bitacora));
                echo anchor(site_url('archivos/index/'.$contrato->Numero), img(array('src' => 'img/archivos.png', 'title' => 'Administrar archivos', 'width' => '30', 'height' => '30')));
                echo anchor(site_url('informes/acta_inicio/'.$contrato->Pk_Id_Contrato), img($acta_inicio));
                ?>
            </td>
        </tr>
    </table>
<?php } ?>