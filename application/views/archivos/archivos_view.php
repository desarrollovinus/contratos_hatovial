<div id="form" class="container_12">
    <div class="grid_11">
        <?php
        if ($this->session->userdata('Tipo') == 1) {
            echo form_open_multipart('archivos/subir/'.$numero_contrato, array('name' => 'f1')); ?>
            <table width="100%">
                <?php
                $categorias = $this->contrato_model->listar_categorias();
                $_categoria = array('' => '');
                
                //Recorrido para traer todas las categorias de la base de datos y agregarlas en el dropdown
                foreach ($categorias as $categoria):
                    $_categoria[$categoria->Pk_Id_Categoria] = $categoria->Categoria;
                endforeach;
                ?>
                <tr>
                    <td style="vertical-align: middle"><?php echo form_dropdown('categoria', $_categoria, set_value('categoria'), 'id="categoria" onchange="cambiar_serie()"'); ?></td>
                        
                    <td style="vertical-align: middle">
                        <select name="subcategoria" id="subcategoria">
                        </select>
                    </td>
                    <td style="vertical-align: middle"><?php echo form_input(array('type' => 'file', 'name' => 'userfile', 'style' => 'width: 400px')); ?></td>
                    <td style="vertical-align: middle"><?php echo form_submit(array('id' => 'guardar', 'name' => 'guardar', 'style' => 'vertical-align: middle'), 'Subir'); ?></td>
                </tr>
            </table>
            <?php
            echo form_close();
            echo form_fieldset('', array('id' => 'fieldset', 'class' => 'fieldset', 'style' => 'text-align:center;'));
            echo form_fieldset_close();
        }
        ?>

        <!--Tabla con los archivos del contrato-->
        <table cellpadding="0" cellspacing="0" border="" class="display" id="example" style="font-size: 15px;">
            <thead>
                <tr>
                    <th class="primero" style="text-align: center">Archivo</th>
                    <th style="text-align: center;">Tama&ntilde;o</th>
                    <th class="ultimo" style="text-align: center; width: 5%">Opciones</th>
                </tr>
            </thead>
            <tbody>
                <?php
                //Se hace el recorrido de los archivos
                foreach ($archivos as $archivo):
                    //Se consulta el tamano del archivo (en bytes)
                    $tamano = filesize($carpeta.'/'.$archivo);
                    
                    //Se valida el tamano y se imprime en KB o MB. 
                    if($tamano > 999999){
                        $tamano = number_format(filesize($carpeta.'/'.$archivo)/1024000, 2, ',', '.').' MB';
                    }else{
                        $tamano = number_format(filesize($carpeta.'/'.$archivo)/1024, 0, ',', '.').' KB';
                    }
                ?>
                <tr>
                    <td><?php echo $archivo; ?></td>
                    <td style="text-align: right"><?php echo $tamano;  ?></td>
                    <td><?php echo anchor_popup(base_url().$carpeta.'/'.$archivo, img(array('src' => 'img/pdf.png', 'title' => 'Abrir archivo', 'width' => '30', 'heihgt' => '30')), array('width' => '800','height' => '600','scrollbars' => 'yes','status' => 'yes','resizable' => 'yes','screenx' => '0','screeny' => '0')); ?></td>
                </tr>
                <?php
                endforeach;
                ?>
            </tbody>
        </table>
    </div>
</div>
<script type='text/javascript'>
    $(document).ready(function() {
        //Cuando se seleccione una categoria, cargara las subcategorias
        $("#categoria").change(function() {
            $("#categoria option:selected").each(function() {
                //se declara la variable con el id de la categoria
                categoria = $('#categoria').val();
                
                //Se envia la variable por post
                $.post("<?php echo site_url('archivos/subcategorias'); ?>", {
                    categoria : categoria
                }, function(data) {
                    //Se cargan en el campo ciudad
                    $("#subcategoria").html(data);
                });
            });
        });
    
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