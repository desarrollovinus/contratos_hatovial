<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de bit&aacute;cora.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Bitacora extends CI_Controller{
    /**
    * Funci&oacute;n constructora de la clase bitacora. 
    * 
    * Se hereda el mismo constructor de la clase Controller para evitar sobreescribirlo y de esa manera 
    * conservar el funcionamiento de controlador.
    * 
    * @access	public
    */
   function __construct() {
        //con esta linea se hereda el constructor de la clase Controller
        parent::__construct();
        //Verifica si el usuario est&aacute; logueado y no pueda acceder por url a menus sin permisos
        if($this->session->userdata('Pk_Id_Usuario') != TRUE || $this->session->userdata('Tipo') == FALSE)
        {
            //Se redirecciona para que cierre la sesion y lo lleve al inicio
            redirect('sesion/cerrar_sesion');
        }
        //Se carga el modelo de la bit&aacute;cora en la base de datos
        $this->load->model('bitacora_model');
   }//Fin construct()
   
    /**
    * Muestra la vista principal del m&oacute;dulo de bit&aacute;cora.
    * 
    * @access	public
    */
    function index(){
        $id_contrato = $this->uri->segment(3);
        $this->data['id_contrato'] = $this->uri->segment(3);
        $this->data['bitacoras'] = $this->bitacora_model->listar_bitacora($id_contrato);
        
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Bit&aacute;cora';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'bitacora/bitacora_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
   }//Fin index()
   
   /**
    * Funci&oacute;n que se encarga de agregar una anotaci&oacute;n a la
    * bit&aacute;cora del contrato espec&iacute;fico
    * 
    * @access	public
    */
   function agregar(){
       /*
        * Estas son las validaciones que se realizan con el Helper form_validation
        */
        $this->form_validation->set_rules('asunto', 'El asunto', 'required|trim');
       
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
       
        /*
         * Esta es la condición que ejecuta las reglas y no lo deja pasar. 
         * Si el método devuelve FALSE, la validación no se llevó corretamente
         */
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Hacen falta algunos datos necesarios<br> para guardar el contrato. Verifique por favor";
            $this->index();
        }else{
            /*
             * Se preparan los datos del contrato para enviar a la base de datos:
             * 
             */
            $bitacora = array(
                'Fecha' => date('Y-m-d H:i:s', time()),
                'Asunto' => $this->input->post('asunto'),
                'Observacion' => $this->input->post('observacion'),
                'Fk_Id_Contratos' => $this->input->post('id_contrato'),
                'Fk_Id_Usuario' => $this->session->userdata('Pk_Id_Usuario')
            );
            
            //Se obtienen los ids para la gesti&oacute;n de auditor&iacute;a
            $numero = $this->input->post('numero');
            
            //Estos datos se insertan en la tabla contratos mediante este modelo
            $this->bitacora_model->agregar_bitacora($bitacora);
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->agregar_bitacora($numero);
            
            //Se establece el mensaje de exito
            $this->data['mensaje_exito'] = 'El registro en la bitácora ha sido creado con &eacute;xito';
            
            /*
             * Se prepara la redirecci&oacute;n
             */
            $id_contrato = $this->uri->segment(3);
            $this->data['id_contrato'] = $this->uri->segment(3);
            $this->data['bitacoras'] = $this->bitacora_model->listar_bitacora($id_contrato);
        
            //se establece el titulo de la p&aacute;gina
            $this->data['titulo'] = 'Contratos - Bit&aacute;cora';
            //se establece la vista que tiene el contenido principal
            $this->data['contenido_principal'] = 'bitacora/bitacora_view';
            //se carga el template
            $this->load->view('includes/template', $this->data);
       }
   }//Fin agregar()
}//Fin bitacora
/* End of file bitacora.php */
/* Location: ./contratos/application/controllers/bitacora.php */
?>
