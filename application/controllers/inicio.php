<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de Inicio.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Inicio extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase Inicio. Esta funci&oacute;n se encarga de verificar que se haya
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
        //Se carga el modelo de los contratos en la base de datos
        $this->load->model('contrato_model');
        //Se carga el helper html para usar en la vista
        $this->load->helper('html');
        //
        $this->load->library('encrypt');
    }//Fin construct()
    
    /**
     * Muestra la vista principal de la aplicaci&oacute;n.
     * 
     * @access	public
     */
    function index(){
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
    }//Fin index()
}//Fin Inicio
/* End of file Inicio.php */
/* Location: ./contratos/application/controllers/Inicio.php */
