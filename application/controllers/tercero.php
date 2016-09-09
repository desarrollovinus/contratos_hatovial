<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo Terceros.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Tercero extends CI_Controller{
    /**
    * Funci&oacute;n constructora de la clase conratista. 
    * 
    * Se hereda el mismo constructor de la clase Controller para evitar sobreescribirlo y de esa manera 
    * conservar el funcionamiento de controlador.
    * 
    * @access	public
    */
    function __construct() {
        //con esta linea se hereda el constructor de la clase Controller
        parent::__construct();
        ///Verifica si el usuario est&aacute; logueado y no pueda acceder por url a menus sin permisos
        if($this->session->userdata('Pk_Id_Usuario') != TRUE || $this->session->userdata('Tipo') == FALSE)
        {
            //Se redirecciona para que cierre la sesion y lo lleve al inicio
            redirect('sesion/cerrar_sesion');
        }
        //Se carga el helper form, el cual servir&aacute; para armar el formulario
        $this->load->helper('form');
        //Se carga el helper html para usar en la vista
        $this->load->helper('html');
        //Se cargan los modelos
        $this->load->model('tercero_model');
        $this->load->model('contrato_model');
    }//Fin construct()
    
    /**
     * Muestra la vista principal del m&oacute;dulo terceros
     * 
     * @access	public
     */
    function index(){
        //Se cargan los tipos de terceros
        $this->data['terceros'] = $this->tercero_model->listar_tipo_terceros();
        //print_r($this->tercero_model->listar_tipo_terceros());
        $this->data['titulo'] = 'Contratos - Terceros';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'tercero/tercero_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
    
    /**
     * Muestra la vista principal del m&oacute;dulo terceros
     * 
     * @access	public
     */
    function agregar(){
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules('tipo', 'El tipo de tercero', 'required|trim');
        $this->form_validation->set_rules('nombre', 'El nombre del tercero', 'required|trim');
        $this->form_validation->set_rules('documento');
        $this->form_validation->set_rules('telefono', 'El n&uacute;mero de tel&eacute;fono', 'trim');
        $this->form_validation->set_rules('direccion');
        $this->form_validation->set_rules('representante_legal');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', '%s no puede estar vac&iacute;o');
        $this->form_validation->set_message('numeric', '%s no puede contener letras');
        
        /*
         * Esta es la condición que ejecuta las reglas y no lo deja pasar. 
         * Si el método devuelve FALSE, la validación no se llevó corretamente
         */
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Faltan datos obligatorios por llenar";
            $this->index();
        }else{
            /*
             * Se preparan los datos del contratista para enviar a la base de datos:
             * 
             */
            $tercero = array(
                'Fk_Id_Terceros_Tipo' => $this->input->post('tipo'),
                'Documento' => $this->input->post('documento'),
                'Nombre' => $this->input->post('nombre'),
                'Telefono' => $this->input->post('telefono'),
                'Direccion' => $this->input->post('direccion'),
                'Representante_Legal' => $this->input->post('representante_legal')
            );
            
            //Impresi&oacute;n del array para verificar que lleguen bien los datos
            //print_r($tercero);
            
            //Se envian los datos del contratista al modelo
            $this->tercero_model->registrar_tercero($tercero);
            
            //Despues de guardado el tercero, obtenemos el id generado
            $id_tercero = $this->db->insert_id();
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->insertar_tercero($id_tercero, $this->input->post('nombre'));
            
            //Se establece el mensaje de exito
            $this->data['mensaje_exito'] = 'Se ha creado "'.$this->input->post('nombre').'"  correctamente';
            
            /*
             * Se prepara la redirecci&oacute;n
             */
            $id_contrato = '';
            //Se traen todos los contratos
            $this->data['contratos'] = $this->contrato_model->listar_contratos($id_contrato);
            //Se traen los estados de los contratos
            $this->data['contratos_estados'] = $this->contrato_model->listar_contratos_estados();
            //se establece el titulo de la p&aacute;gina
            $this->data['titulo'] = 'Contratos - Inicio';
            //se establece la vista que tiene el contenido principal
            $this->data['contenido_principal'] = 'inicio/inicio_view';
            //se carga el template
            $this->load->view('includes/template', $this->data);
        }
    }//Fin agregar()
    
    /**
     * Visualiza todos los terceros registrados
     * 
     * @access	public
     */
    function ver(){
        //Se listan los terceros
        $this->data['terceros'] = $this->tercero_model->listar_terceros(NULL);
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Terceros';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'tercero/ver_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin ver()
    
    /**
     * Permite editar un tercero
     * 
     * @access	public
     */
    function actualizar(){
        //Se toma el id del tercero
        $this->data['id_tercero'] = $this->input->post('id_tercero');
        $id_tercero = $this->input->post('id_tercero');
        //Se listan los terceros
        $this->data['terceros'] = $this->tercero_model->listar_terceros($id_tercero);
        //Se cargan los tipos de terceros
        $this->data['tipo_terceros'] = $this->tercero_model->listar_tipo_terceros();
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Modificar tercero';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'tercero/actualizar_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin actualizar
    
    /**
     * Permite editar un tercero
     * 
     * @access	public
     */
    function actualizar_tercero(){
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules('tipo', 'El tipo de tercero', 'required|trim');
        $this->form_validation->set_rules('nombre', 'El nombre del tercero', 'required|trim');
        $this->form_validation->set_rules('documento');
        $this->form_validation->set_rules('telefono', 'El n&uacute;mero de tel&eacute;fono', 'trim');
        $this->form_validation->set_rules('direccion');
        $this->form_validation->set_rules('representante_legal');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', '%s no puede estar vac&iacute;o');
        $this->form_validation->set_message('numeric', '%s no puede contener letras');
        
        /*
         * Esta es la condición que ejecuta las reglas y no lo deja pasar. 
         * Si el método devuelve FALSE, la validación no se llevó corretamente
         */
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Faltan datos obligatorios por llenar";
            $this->actualizar();
        }else{
            /*
             * Se preparan los datos del contratista para enviar a la base de datos:
             * 
             */
            $tercero = array(
                'Fk_Id_Terceros_Tipo' => $this->input->post('tipo'),
                'Documento' => $this->input->post('documento'),
                'Nombre' => $this->input->post('nombre'),
                'Telefono' => $this->input->post('telefono'),
                'Direccion' => $this->input->post('direccion'),
                'Representante_Legal' => $this->input->post('representante_legal')
            );
            //print_r($tercero);
            
            //Se ejecuta el modelo que actualiza los datos
            $this->tercero_model->actualizar_tercero($tercero, $this->input->post('id_tercero'));
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->modificar_tercero($this->input->post('id_tercero'), $this->input->post('nombre'));
            
            //Se establece el mensaje de exito
            $this->data['mensaje_exito'] = 'El tercero '.$this->input->post('nombre').' ha sido actualizado con &eacute;xito';
            
            //Se listan los terceros
            $this->data['terceros'] = $this->tercero_model->listar_terceros(NULL);
            //se establece el titulo de la p&aacute;gina
            $this->data['titulo'] = 'Contratos - Terceros';
            //se establece la vista que tiene el contenido principal
            $this->data['contenido_principal'] = 'tercero/ver_view';
            //se carga el template
            $this->load->view('includes/template', $this->data);
        }
    }//Fin actualizar_tercero
}//Fin tercero
/* End of file tercero.php */
/* Location: ./contratos/application/controllers/tercero.php */
