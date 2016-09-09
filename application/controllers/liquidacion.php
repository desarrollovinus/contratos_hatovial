<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de liquidaci&oacute;n.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Liquidacion extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase liquidacion. 
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
        //Se carga el modelo de los contratos en estado de liquidaci&oacute;n en la base de datos
        $this->load->model('liquidacion_model');
        //Se carga el helper html para usar en la vista
        $this->load->helper('html');
    }//Fin construct()
    
    /**
     * Muestra la vista principal del m&oacute;dulo de liquidaci&oacute;n.
     * 
     * @access	public
     */
    function index(){
        $this->data['contratos'] = $this->liquidacion_model->listar_contratos();
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Liquidaciones';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'liquidacion/liquidacion_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
}//Fin liquidacion
/* End of file liquidacion.php */
/* Location: ./contratos/application/controllers/liquidacion.php */
?>
