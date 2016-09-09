<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo Archivos
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Archivos extends CI_Controller{
    /**
    * Funci&oacute;n constructora de la clase archivos. 
    * 
    * Se hereda el mismo constructor de la clase Controller para evitar sobreescribirlo y de esa manera 
    * conservar el funcionamiento de controlador.
    * 
    * @access	public
    */
    function __construct() {
        //con esta linea se hereda el constructor de la clase Controller
        parent::__construct();
        if($this->session->userdata('Pk_Id_Usuario') != TRUE)
        {
            //Redirecci&oacute;n al controlador de sesi&oacute;n
            redirect('sesion');
        }
        //Se cargan los modelos y helpers
        $this->load->model('contrato_model');
        $this->load->helper('html');
    }//Fin construct()
    
    var $ruta = "archivos/";
    
    /**
     * Muestra la vista principal del m&oacute;dulo Actualizar.
     * 
     * @access	public
     */
    function index(){
        //Se obtiene el numero del contrato
        $this->data['numero_contrato'] = $this->uri->segment(3);
        
        //Valida que la carpeta exista. Si no existe, la crea
        if( ! is_dir($this->ruta.$this->data['numero_contrato'])){
            @mkdir($this->ruta.$this->data['numero_contrato'], 0777);
        }
        
        //Carga los archivos pertenecientes a ese contrato
        $this->configuracion($this->data['numero_contrato']);
            
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Archivos';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'archivos/archivos_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
    
    /*
     * Funcion encargada de subir el archivo,
     * dependiendo los parametros que se env&iacute;en
     */
    function subir(){
        //Se obtiene el numero de contrato
        $numero_contrato = $this->uri->segment(3);
        
        //Se obtiene la serie para nombrar el archivo
        $serie = $this->input->post('subcategoria');
        
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules('subcategoria', 'Debe seleccionar un tipo de documento y una serie para subir', 'required');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', '');
        
        /*
         * Esta es la condición que ejecuta las reglas y no lo deja pasar. 
         * Si el método devuelve FALSE, la validación no se llevó corretamente
         */
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Debe seleccionar un tipo de documento y una serie";
            $this->index();
        }else{
            //Tipo de archivo
            if($serie == '16 Otrosi' or $serie == '14 Acta Obra' or $serie == '15 Factura'){
                $config['overwrite'] = FALSE;                                   //No Sobreescribe el archivo si ya existe
            }else{
                $config['overwrite'] = TRUE;                                   //Sobreescribe el archivo si ya existe
            }

            //Se realiza la configuraci&oacute;n
            $config['upload_path'] = $this->ruta.$numero_contrato;              //Ruta donde se va a guardar el archivo
            $config['file_name'] = $serie.' '.$numero_contrato.'.pdf';          //Nombre del archivo. Se guarda con un nombre espec&iacute;fico
            $config['allowed_types'] = 'pdf';                                   //Formatos de archivo permitidos
            $config['remove_spaces'] = TRUE;                                    //Convierte los espacios en blanco en guiones bajos
            $config['max_size'] = 0;                                            //limite de tamano para el archivo. 0 = Sin limite

            //
            $nombre = $serie.' '.$numero_contrato.'.pdf';
            
            //Se carga la librer&iacute;a
            $this->load->library('upload', $config);

            //Valida y realiza la subida
            if (!  $this->upload->do_upload()){
                //Se escribe el mensaje de error
                $this->data['mensaje_error'] = $this->upload->display_errors();

                //Se carga la configuraci&oacute;n
                $this->configuracion($numero_contrato);

                //Se obtiene el numero del contrato
                $this->data['numero_contrato'] = $numero_contrato;
                //se establece el titulo de la p&aacute;gina
                $this->data['titulo'] = 'Contratos - Archivos';
                //se establece la vista que tiene el contenido principal
                $this->data['contenido_principal'] = 'archivos/archivos_view';
                //se carga el template
                $this->load->view('includes/template', $this->data);
            }else{
                //Se realiza la subida
                array('upload_data' => $this->upload->data());

                //Se escribe el mensaje de exito
                $this->data['mensaje_exito'] = "El archivo se ha subido correctamente";

                //Se carga la configuraci&oacute;n
                $this->configuracion($numero_contrato);
                
                //Se realiza la acci&oacute; de auditor&iacute;a
                $this->auditoria_model->subir_archivo($nombre);

                //Se obtiene el numero del contrato
                $this->data['numero_contrato'] = $numero_contrato;
                //se establece el titulo de la p&aacute;gina
                $this->data['titulo'] = 'Contratos - Archivos';
                //se establece la vista que tiene el contenido principal
                $this->data['contenido_principal'] = 'archivos/archivos_view';
                //se carga el template
                $this->load->view('includes/template', $this->data);
            }
        }
    }//Fin subir()
    
    /*
     * Lista los archivos existentes para un contrato espec&iacute;fico y los env&iacute;a
     * a la vista para que los muestre
     */
    function configuracion($numero_contrato){
        //Se cargan las p&oacute;lizas
        $this->data['polizas'] = $this->contrato_model->listar_polizas();
        
        //Abre el directorio
        if($carpeta = opendir($this->ruta.$numero_contrato)){
            //Se construye un arreglo con los nombres de archivo
            $nombres = array();
            
            //Recorre los archivos que tiene
            while(($archivo = readdir($carpeta)) !== FALSE){
                if($archivo != '.' && $archivo != '..'){
                    //se guardan los nombres en el array
                    array_push($nombres, $archivo);
                }
            }
            
            //se cierra la carpeta
            closedir();
            
            //Se crean las variables para enviar a la vista
            $this->data['archivos'] = $nombres;
            $this->data['carpeta'] = $this->ruta.$numero_contrato;
        }
    }//Fin configuracion()
    
    /*
     * Lista las subcategorias
     */
    function subcategorias(){
        //Se recibe el id de la categoria
        $categoria = $this->input->post('categoria');
        
        //Se cargan las subcategorias con ese id de categoria
        $subcategorias = $this->contrato_model->listar_subcategorias($categoria);
        
        //Se abre el select
        $respuesta .= '<select name="subcategoria">';
        $respuesta .= '<option value=""></option>';
        
        
        //Se imprime cada subcategoria imprimiendo el resultado de la consulta al modelo
        foreach($subcategorias as $subcategoria):
            $respuesta .= '<option value="'.$subcategoria->Subcategoria.'">'.$subcategoria->Subcategoria.'</option>'; 
        endforeach;
        
        //Se cierra el select
        $respuesta .= '</select>';
        
        //Se imprime el select completo
        echo $respuesta;
    }//Fin subcategorias()
}//Fin archivos
/* End of file archivos.php */
/* Location: ./contratos/application/controllers/archivos.php */