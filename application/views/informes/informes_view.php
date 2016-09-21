<?php
//Atributos para el fieldset
$fieldset = array('id' => 'fieldset', 'class' => 'fieldset', 'style' => 'text-align:center;');
$icono = base_url('img/excel.png');
$icono_pdf = base_url('img/icono_pdf.png');
?>

<div id="form" class="container_12">
    <div class="grid_11">
        <center><b>Generador de informes seg&uacute;n el filtro seleccionado</b></center><br/>

        <table width="100%">
            <tr>
                <td>
                    <div id="informes">
                        <table width="100%">
                            <?php
                            echo form_fieldset('Por contratista', $fieldset);
                            echo form_fieldset_close();

                            $_contratista = array('' => 'TODOS');
                            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                            foreach($contratistas as $contratista):
                                $_contratista[$contratista->Pk_Id_Terceros] = $contratista->Nombre;
                            endforeach;
                            ?>
                            <tr>
                                <td><?php echo form_label('Seleccione el contratista', 'id_contratista'); ?></td>
                                <td rowspan="2">
                                    <?php
                                    /*
                                     * Informe Excel - Por contratista
                                     */
                                    echo form_open('informes/contratistas_excel');
                                    // echo form_hidden('id_contratista', $contratista->Pk_Id_Terceros);
                                    echo form_submit(array( 'type' => 'image', 'src' => $icono, 'target' => 'blank'));
                                    ?>
                                </td>
                            </tr>
                            <tr>
                                <td><?php echo form_dropdown('id_contratista', $_contratista); ?></td>
                            </tr>
                            <?php echo form_close(); ?>
                        </table>
                    </div>
                </td>

                <!-- OK -->
                <td>
                    <div id="informes">
                        <table width="100%">
                            <?php
                            echo form_fieldset('Por contratante', $fieldset);
                            echo form_fieldset_close();

                            $_contratante = array('' => 'TODOS');
                            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                            foreach($contratantes as $contratante):
                                $_contratante[$contratante->Pk_Id_Terceros] = $contratante->Nombre;
                            endforeach;
                            ?>
                            <tr>
                                <td><?php echo form_label('Seleccione el contratante', 'id_contratante'); ?></td>
                                <td rowspan="2">
                                    <?php
                                    /*
                                     * Informe Excel - Por contratista
                                     */
                                    echo form_open('informes/contratantes_excel');
                                    // echo form_hidden('id_contratista', $contratista->Pk_Id_Terceros);
                                    echo form_submit(array( 'type' => 'image', 'src' => $icono, 'target' => 'blank'));
                                    ?>
                                </td>
                            </tr>
                            <tr>
                                <td><?php echo form_dropdown('id_contratante', $_contratante); ?></td>
                            </tr>
                            <?php echo form_close(); ?>
                        </table>
                    </div>
                </td>
                <td>
                    <div id="informes">
                        <?php
                        /*
                         * Informe 1 - Por estados
                         */
                        echo form_fieldset('Por estado', $fieldset);
                        echo form_fieldset_close();

                        //Formulario informe 1
                        echo form_open('informes/estados');
                        ?>
                        <table width="100%">
                            <?php
                            $_estado_contrato = array('' => '');
                            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                            foreach($contratos_estados as $estados):
                                $_estado_contrato[$estados->Pk_Id_Estado] = $estados->Estado;
                            endforeach;
                            ?>
                            <tr>
                                <td><?php echo form_label('Seleccione un estado', 'estados_contrato'); ?></td>
                                <td rowspan="2">
                                    <?php
                                    echo form_submit(array( 'type' => 'image', 'src' => $icono));
                                    echo form_close();
                                    ?>
                                </td>
                            </tr>
                            <tr>
                                <td><?php echo form_dropdown('estados_contratos', $_estado_contrato); ?></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <div id="informes">
                        <?php
                        /*
                         * Informe 3 - Contratos sin acta de inicio
                         */
                        echo form_fieldset('Sin acta de inicio', $fieldset);
                        echo form_fieldset_close();

                        //Formulario informe 3
                        echo form_open('informes/no_acta_inicio');
                        ?>
                        <table width="100%">
                            <tr>
                                <td>
                                    <td><?php echo form_submit(array( 'type' => 'image', 'src' => $icono)); ?></td>
                                </td>
                            </tr>
                        </table>
                        <?php echo form_close();?>
                    </div>
                </td>
                <td>
                    <div id="informes">
                        <?php
                        /*
                         * Informe 2 - Por fecha inicial
                         */
                        echo form_fieldset('Por fecha de inicio', $fieldset);
                        echo form_fieldset_close();

                        //Formulario informe 2
                        echo form_open('informes/fecha_inicial');
                        ?>
                        <table width="100%">
                            <tr>
                                <td><?php echo form_input(array('name' => 'fecha1', 'id' => 'fecha1', 'style' => 'width: 140px;', 'readonly' => 'readonly', 'value' => set_value('fecha1'))); ?></td>
                                <td rowspan="2"><?php  echo form_submit(array( 'type' => 'image', 'src' => $icono)); ?></td>
                            </tr>
                            <tr>
                                <td><?php echo form_input(array('name' => 'fecha2', 'id' => 'fecha2', 'style' => 'width: 140px;', 'readonly' => 'readonly', 'value' => set_value('fecha2'))); ?></td>
                            </tr>
                            <tr>
                                <?php echo form_close();?>
                            </tr>
                        </table>
                    </div>
                </td>
                <td>
                    <div id="informes">
                        <?php
                        /*
                         * Informe 3 - Por fecha de vencimiento
                         */
                        echo form_fieldset('Por fecha de vencimiento', $fieldset);
                        echo form_fieldset_close();

                        //Formulario informe 3
                        echo form_open('informes/fecha_vencimiento');
                        ?>
                        <table width="100%">
                            <tr>
                                <td><?php echo form_input(array('name' => 'fecha3', 'id' => 'fecha3', 'style' => 'width: 140px;', 'readonly' => 'readonly', 'value' => set_value('fecha3'))); ?></td>
                                <td rowspan="2"><?php  echo form_submit(array( 'type' => 'image', 'src' => $icono)); ?></td>
                            </tr>
                            <tr>
                                <td><?php echo form_input(array('name' => 'fecha4', 'id' => 'fecha4', 'style' => 'width: 140px;', 'readonly' => 'readonly', 'value' => set_value('fecha4'))); ?></td>
                            </tr>
                            <tr>
                                <?php echo form_close();?>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <div id="informes">
                        <table width="100%">
                            <?php
                            echo form_fieldset('Pagos', $fieldset);
                            echo form_fieldset_close();

                            $_contratista = array('' => 'TODOS');
                            //Recorrido para traer todos los estados de la base de datos y agregarlos en el dropdown
                            foreach($contratistas as $contratista):
                                $_contratista[$contratista->Pk_Id_Terceros] = $contratista->Nombre;
                            endforeach;
                            ?>
                            <tr>
                                <td><?php echo form_label('Seleccione el contratista', 'id_contratista'); ?></td>
                                <td rowspan="2">
                                    <?php
                                    /*
                                     * Informe Excel - Por contratista
                                     */
                                    echo form_open('informes/contratistas_excel');
                                    // echo form_hidden('id_contratista', $contratista->Pk_Id_Terceros);
                                    echo form_submit(array( 'type' => 'image', 'src' => $icono, 'target' => 'blank'));
                                    ?>
                                </td>
                            </tr>
                            <tr>
                                <td><?php echo form_dropdown('id_contratista', $_contratista); ?></td>
                            </tr>
                            <?php echo form_close(); ?>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
    </div>
</div>
