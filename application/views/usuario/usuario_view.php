<?php
/*
 * Arrays que almacenan los atributos de cada campo del formulario
 */
$atributos = array('id' => 'registro_usuario', 'class' => 'id_registro_usuario');
$nombres = array('id' => 'nombres', 'name' => 'nombres', 'value' => set_value('nombres'));
$apellidos = array('id' => 'apellidos', 'name' => 'apellidos', 'value' => set_value('apellidos'));
$usuario = array('id' => 'usuario', 'name' => 'usuario', 'value' => set_value('usuario'));
$password = array('id' => 'password', 'name' => 'password', 'value' => set_value(''));
$repassword = array('id' => 'repassword', 'name' => 'repassword', 'value' => set_value(''));
$email = array('id' => 'email', 'name' => 'email', 'value' => set_value('email'));
$telefono = array('id' => 'telefono', 'name' => 'telefono', 'value' => set_value('telefono'));
$guardar = array('name'      => 'guardar',    'id'        => 'guardar',    'value'     => 'Guardar');

//Se inicia el formulario
echo form_open('usuario/registrar_usuario', $atributos);
?>
<div class="container_12">
    <div class="titulos_formularios">Registro de usuario nuevo</div><br>
    <div id="form" class="grid_12" align="center">
        <table align="center" width="600px">
            <tr>
                <th width="25%"><?php echo form_label('Nombres*'); ?></th>
                <td width="20%"><?php echo form_input($nombres); ?></td>
                <td width="70%">
                    <span class="error">
                        <?php echo form_error('nombres'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <th><?php echo form_label('Apellidos*'); ?></th>
                <td><?php echo form_input($apellidos); ?></td>
                <td>
                    <span class="error">
                        <?php echo form_error('apellidos'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <th><?php echo form_label('Nombre Usuario*'); ?></th>
                <td><?php echo form_input($usuario); ?></td>
                <td>
                    <span class="error">
                        <?php echo form_error('usuario'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <th><?php echo form_label('Clave*'); ?></th>
                <td><?php echo form_password($password); ?></td>
                <td>
                    <span class="error">
                        <?php echo form_error('password'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <th><?php echo form_label('Repita la clave*'); ?></th>
                <td><?php echo form_password($repassword); ?></td>
                <td>
                    <span class="error">
                        <?php echo form_error('repassword'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <th><?php echo form_label('Email'); ?></th>
                <td><?php echo form_input($email); ?></td>
                <td>
                    <span class="error">
                        <?php echo form_error('email'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <th><?php echo form_label('Tel&eacute;fono'); ?></th>
                <td><?php echo form_input($telefono); ?></td>
                <td>
                    <span class="error">
                        <?php echo form_error('telefono'); ?>
                    </span>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <div align="center"><?php
                    echo form_submit($guardar);                    
                    ?></div>
                </td>
            </tr>
        </table>
        <?php
        //Se cierra o finaliza el formulario
        echo form_close();
        ?>
    </div>
</div>