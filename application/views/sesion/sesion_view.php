<?php
//Se inicia el formulario
echo form_open('sesion/validar_login');
?>
<div class="container_12">
    <div class="titulos_formularios">Identificaci&oacute;n</div><br>
    <div id="form" align="center">
        <table width="40%">
            <tr>
                <th><?php echo form_label('Nombre de usuario'); ?></th>
                <td>
                    <div style="alignment-adjust: right">
                        <?php echo form_input(array('name' => 'usuario', 'id' => 'usuario', 'value' => set_value('usuario'))); ?>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="error"><?php echo form_error('usuario'); ?></td>
            </tr>
            <tr>
                <th><?php echo form_label('Contrase&ntilde;a&nbsp;'); ?></th>
                <td>
                    <div>
                        <?php echo form_password(array('name' => 'password', 'id' => 'password', 'value' => '')); ?>
                    </div>
                </td>
            </tr>
            <tr class="error">
                <td colspan="2"><?php echo form_error('password'); ?></td>
            </tr>
            <tr style="alignment-adjust: center">  
                <td colspan="2">
                    <center><?php echo form_submit(array('name' => 'submit', 'id' => 'submit', 'value' => 'Ingresar')); ?></center><br>
                </td>
            </tr>
        </table>
        <?php
        //Se cierra o finaliza el formulario
        echo form_close();
        ?>
    </div>
</div>
<script type="text/javascript">
    //Cuando el DOM este listo
    $(document).ready(function(){
        //Se establece el foco del campo
        $("#usuario").focus();
    });//Fin document.ready
</script>

