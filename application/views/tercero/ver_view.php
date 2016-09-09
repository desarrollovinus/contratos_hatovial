<div id="form" class="container_12">
    <div class="grid_11">
        <div class="titulos_formularios">Listado de terceros</div><br>
        <table cellpadding="0" cellspacing="0" border="" class="display" id="example" style="font-size: 15px;">
            <thead>
                <tr>
                    <th class="primero">Tipo</th>
                    <th>Nombre</th>
                    <th>Documento</th>
                    <th>Tel&eacute;fono</th>
                    <th>Direcci&oacute;n</th>
                    <th>Representante legal</th>
                    <th class="ultimo">Editar</th>
                </tr>
            </thead>
            <tbody>
                <?php foreach ($terceros as $tercero): ?>
                <tr>
                    <td><?php echo $tercero->Tipo; ?></td>
                    <td><?php echo $tercero->Nombre; ?></td>
                    <td><?php echo $tercero->Documento; ?></td>
                    <td><?php echo $tercero->Telefono; ?></td>
                    <td><?php echo $tercero->Direccion; ?></td>
                    <td><?php echo $tercero->Representante_Legal; ?></td>
                    <td>
                        <?php
                        echo form_open('tercero/actualizar');
                        echo form_hidden('id_tercero', $tercero->Pk_Id_Terceros);
                        echo form_submit(array('type' => 'image', 'src' => base_url().'img/modificar.png', 'width' => '25', 'height' => '25'));
                        echo form_close();//echo anchor(site_url('tercero/actualizar/'.), img(array('src' => 'img/modificar.png', 'title' => 'Ver contrato', 'width' => '20', 'height' => '20'))) ?>
                    </td>
                </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
    </div>
</div>

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