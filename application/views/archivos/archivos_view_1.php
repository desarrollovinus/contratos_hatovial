<div id="form" class="container_12">
    <div class="grid_11">
        <?php echo form_open_multipart('archivos/subir/'.$numero_contrato, array('name' => 'f1')); ?>
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
                    <!--Series-->
                    <select name = serie> 
                        <option value=""></option>
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
    //Contrato
    var series_1 = new Array(
        "",
        "01 Orden",
        "05 Contrato",
        "06 Anexo Contrato",
        "07 Acta Inicio",
        "16 Solicitud Modificacion",
        "17 Acta de Recibo",
        "18 Ata Liquidacion-Documentos",
        "19 Acta Suspension"
    );
    
    //Contratista
    var series_2 = new Array(
        "",
        "02 Certificado Existencia Representacion Legal",
        "03 Copia CC Representante Legal",
        "04 Copia RUT Representante Legal"
    ) ;
    
    //Polizas
    var series_3 = new Array(
        "",
        "08 Poliza Cumplimiento",
        "09 Poliza Prestaciones Sociales",
        "10 Poliza Anticipos",
        "11 Poliza Calidad",
        "12 Poliza Estabilidad",
        "13 Poliza RC"
    );
    
    //Actas de obra o facturas
    var series_4 = new Array(
        "",
        "14 Acta Obra",
        "15 Factura"
    );
    
    //Otrosi
    var series_5 = new Array(
        "",
        "20 Otrosi Poliza Cumplimiento",
        "20 Otrosi Poliza Prestaciones Sociales",
        "20 Otrosi Poliza Anticipos",
        "20 Otrosi Poliza Calidad",
        "20 Otrosi Poliza Estabilidad",
        "20 Otrosi Poliza RC"
    );
    
    //Acta de reinicio
    var series_6 = new Array(
        "",
        "21 Reinicio Poliza Cumplimiento",
        "21 Reinicio Poliza Prestaciones Sociales",
        "21 Reinicio Poliza Anticipos",
        "21 Reinicio Poliza Calidad",
        "21 Reinicio Poliza Estabilidad",
        "21 Reinicio Poliza RC"
    );
   
    /*
     * Esta funcion
     */
    function cambiar_serie(){
        //Tomo el valor del select de la categoria elegida
        var categoria 
        categoria = document.f1.categoria[document.f1.categoria.selectedIndex].value 
        //miro si la categoria está definida
        if (categoria != 0) { 
         //si estaba definida, entonces pongo las opciones de la serie correspondiente. 
         //selecciono el array de serie adecuado 
         mis_series=eval("series_" + categoria) 
         //calculo el numero de series 
         numero_series = mis_series.length 
         //marco el número de series en el select 
         document.f1.serie.length = numero_series 
         //para cada serie del array, la introduzco en el select 
            for(i=0; i<numero_series; i++){
                document.f1.serie.options[i].value=mis_series[i] 
                document.f1.serie.options[i].text=mis_series[i] 
            }//Fin for
        }else{
            //si no había serie seleccionada, elimino las series del select 
            document.f1.serie.length = 1 
            //coloco un guión en la única opción que se ha dejado 
            document.f1.serie.options[0].value = "" 
            document.f1.serie.options[0].text = "" 
        } 
            //marco como seleccionada la opción primera de serie 
            document.f1.serie.options[0].selected = true 
    }//Fin cambiar_serie
        
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