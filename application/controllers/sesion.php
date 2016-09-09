<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de la sesi&oacute;n.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Sesion extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase Sesion. 
     * Esta funci&oacute;n se encarga de
    * iniciar la sesi&oacute;n.
    * 
    * Se hereda el mismo constructor de la clase Controller para evitar sobreescribirlo y de esa manera 
    * conservar el funcionamiento de controlador.
    * 
    * @access	public
    */
    function __construct() {
        //con esta linea se hereda el constructor de la clase Controller
        parent::__construct();
        //Se carga el modelo que hace la consulta del usuario en la base de datos
        $this->load->model('usuario_model');
    }//Fin construct()
    
    /**
     * Muestra la vista principal del m&oacute;dulo de sesi&oacute;n.
     * 
     * @access	public
     */
    function index(){
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Identificarse';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'sesion/sesion_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
    
    /**
    * Funci&oacute;n que se encarga de iniciar la sesi&oacute;n siempre y cuando supere la validaci&oacute;n
    * 
    * @access	public
    */
    function validar_login(){
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules('usuario', 'nombre de usuario', 'required');
        $this->form_validation->set_rules('password', 'contrase&ntilde;a', 'required|min_length[4]');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', 'Este campo es obligatorio');
        $this->form_validation->set_message('min_length', 'El campo %s debe tener como m&iacute;nimo 4 caracteres');
        
        /*
         * Esta es la condición que ejecuta las reglas y no lo deja pasar. 
         * Si el método devuelve FALSE, la validación no se llevó corretamente
         */
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de alerta
            $this->data['mensaje_alerta'] = "Faltan datos por llenar";
            $this->index();
        }else{
            //Se traen los datos de la vista
            $usuario = $this->input->post('usuario');
            $password = md5($this->input->post('password'));
           
            //Se usa el modelo para validar los datos ingresados
            $datos_usuario = $this->usuario_model->validar_login($usuario, $password);
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->intento_inicio_sesion($usuario);
                       
            //Se verifica si la consulta trajo alg&uacute;n resultado
            if($datos_usuario){
                //Si trajo resultado, valida que el usuario est&eacute; activo
                if($datos_usuario->Activo != true){
                    //Se imprime el mensaje de error
                    $this->data['mensaje_error'] = "Este usuario est&aacute; desactivado. Verifique con el administrador del sistema por favor.";
                    
                    //Se redirecciona a la misma vista
                    $this->index();
                }else{
                    //Se obtiene el id del usuario para la auditor&iacute;a
                    $id_usuario = $datos_usuario->Pk_Id_Usuario;
                    //Se arma un array indicando los datos que se van a cargan a la sesi&oacute;n
                    $datos_sesion = array(
                        'Pk_Id_Usuario' => $datos_usuario->Pk_Id_Usuario,
                        'Nombres' => $datos_usuario->Nombres,
                        'Apellidos' => $datos_usuario->Apellidos,
                        'Usuario' => $datos_usuario->Usuario,
                        'Email' => $datos_usuario->Email,
                        'Tipo' => $datos_usuario->Tipo
                    );
                    //Se cargan los datos a la sesi&oacute;n
                    $this->session->set_userdata($datos_sesion);
                    
                    //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
                    $this->auditoria_model->iniciar_sesion();
                    
                    //Se redirecciona al contrlador de inicio
                    redirect('');}
                }else{
                    //Se imprime el mensaje de error
                    $this->data['mensaje_error'] = "El usuario y la contrase&ntilde;a no coinciden. Verifique nuevamente por favor.";
                    //Se redirecciona a la misma vista
                    $this->index();
                }
        }
    }//Fin validar_login()
    
    /**
    * Funcion encargada de destruir los datos de la sesi&oacute;n activa.
    * 
    * @access	public
    */
    function cerrar_sesion(){
        //Se traen los datos de la vista para la acci&oacute;nn de auditor&iacute;a
        //$usuario = $this->session->userdata('Usuario');
        //$id_usuario = $this->session->userdata('Pk_Id_Usuario');
        
        //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
        $this->auditoria_model->cerrar_sesion();
        
        //Se destruye la sesion
        $this->session->sess_destroy();
        
        //Se redirecciona al controlador de inicio
        redirect('');
    }//Fin cerrar_sesion()
}//Fin Sesion
/* End of file sesion.php */
/* Location: ./contratos/application/controllers/sesion.php */
?>