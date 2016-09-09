<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de auditor&iacute;a.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Auditoria extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase auditoria. 
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
        //Se carga el helper html para usar en la vista
        $this->load->helper('html');
    }//Fin construct
    
    /**
     * Muestra la vista principal del m&oacute;dulo de auditor&iacute;a.
     * 
     * @access	public
     */
    function index(){
        $this->data['auditorias'] = $this->auditoria_model->listar_auditoria();
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Gesti&oacute;n de auditor&iacute;a';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'auditoria/auditoria_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
}//Fin auditoria
/* End of file auditoria.php */
/* Location: ./contratos/application/controllers/auditoria.php */
