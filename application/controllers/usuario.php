<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de Usuario.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Usuario extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase usuario. Esta funci&oacute;n se encarga de verificar que se haya
    * iniciado sesi&oacute;n, si no se ha iniciado sesi&oacute;n inmediatamente redirecciona hacia Sesion.php.
    * 
    * Se hereda el mismo constructor de la clase Controller para evitar sobreescribirlo y de esa manera 
    * conservar el funcionamiento de controlador.
    * 
    * @access	public
    */
    function __construct() {
        //con esta linea se hereda el constructor de la clase Controller
        parent::__construct();
        //Verifica si el usuario est&aacute; logueado
        if($this->session->userdata('Pk_Id_Usuario') != TRUE)
        {
            //Redirecci&oacute;n al controlador de sesi&oacute;n
            redirect('sesion');
        }
        //Se carga el helper form, el cual servir&aacute; para armar el formulario
        $this->load->helper('form');
        //Se carga form_validation para validar los campos
        $this->load->library('form_validation');
        //Se carga el modelo que hace la consulta del usuario en la base de datos
        $this->load->model('usuario_model');
    }//Fin construct()
    
    /**
     * Muestra la vista principal del m&oacute;dulo de usuario.
     * 
     * @access	public
     */
    function index(){
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Registrarse';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'usuario/usuario_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
    
    /**
     * Valida los nombres de usuario registrados en la aplicaci&oacute;n
     * con el fin de no permitir que se  repitan.
     * 
     * @access	private
     */
    function _validar_usuario($usuario){
        return $this->usuario_model->validar_usuario($usuario);
    }//Fin _validar_usuario()
    
    /**
     * Valida los correos electr&oacute;nicos registrados en la aplicaci&oacute;n
     * con el fin de no permitir que se  repitan.
     * 
     * @access	private
     */
    function _validar_email($email){
        return $this->usuario_model->validar_email($email);
    }//Fin _validar_email()
    
    /**
     * Funci&oacute;n que se encarga de registrar un usuario nuevo
     * en la aplicaci&oacute;n
     * 
     * @access	public
     */
    function registrar_usuario(){
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules('nombres', 'nombres', 'required|trim');
        $this->form_validation->set_rules('apellidos', 'apellidos', 'required|trim');
        $this->form_validation->set_rules('usuario', 'nombre de usuario', 'required|trim|alphanumeric|callback__validar_usuario');
        $this->form_validation->set_rules('_validar_usuario', 'usuario', 'trim|numeric');
        $this->form_validation->set_rules('password', 'clave', 'required|min_length[4]|md5');
        $this->form_validation->set_rules('repassword', 'clave', 'required|matches[password]|min_length[4]|md5');
        $this->form_validation->set_rules('email', 'correo electr&oacute;nioco', 'trim|valid_email|callback__validar_email');
        $this->form_validation->set_rules('_validar_email', 'email', 'trim|numeric');
        $this->form_validation->set_rules('telefono', 'Cedula', 'trim|numeric');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', 'Este campo es obligatorio');
        $this->form_validation->set_message('min_length', 'La %s debe tener como m&iacute;nimo 4 d&iacute;gitos');
        $this->form_validation->set_message('max_length', 'La %s no puede exceder los 12 caracteres');
        $this->form_validation->set_message('matches', 'Las contrase&ntilde;as no coinciden');
        $this->form_validation->set_message('_validar_usuario', 'Este nombre de usuario ya existe. Registre otro');
        $this->form_validation->set_message('valid_email', 'El %s no es v&aacute;lido');
        $this->form_validation->set_message('_validar_email', 'Este correo electr&oacute;nico ya existe. Registre otro');
        $this->form_validation->set_message('numeric', 'En este campo solo se permiten n&uacute;meros');
        
        /*
         * Esta es la condición que ejecuta las reglas y no lo deja pasar. 
         * Si el método devuelve FALSE, la validación no se llevó corretamente
         */
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Faltan datos obligatorios por llenar";
            $this->index();
        }else{
            //Se traen los datos de la vista para la acci&oacute;nn de auditor&iacute;a
            $usuario = $this->input->post('usuario');
            $id_usuario = $this->session->userdata('Pk_Id_Usuario');
            
            /*
             * Se prepara el usuario para enviar a la base de datos
             */
            $usuario_nuevo = array(
                'Nombres' => $this->input->post('nombres'),
                'Apellidos' => $this->input->post('apellidos'),
                'Usuario' => $this->input->post('usuario'),
                'Clave' => $this->input->post('password'),
                'Email' => $this->input->post('email'),
                'Tipo' => $this->input->post('Tipo')
            );
            
            //Se inserta el usuario a la base de datos
            $this->usuario_model->insertar_usuario($usuario_nuevo);
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->insertar_usuario($usuario, $id_usuario);
            
            //Se redirecciona al controlador de inicio
            redirect('');
        }
    }//Fin registrar_usuario
}//Fin usuario
/* End of file usuario.php */
/* Location: ./contratos/application/controllers/usuario.php */