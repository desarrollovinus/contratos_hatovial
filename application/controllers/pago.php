<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de Pagos.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Pago extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase pago. Esta funci&oacute;n se encarga de verificar que se haya
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
        ///Verifica si el usuario est&aacute; logueado
        if($this->session->userdata('Pk_Id_Usuario') != TRUE)
        {
            //Se redirecciona para que cierre la sesion y lo lleve al inicio
            redirect('sesion/cerrar_sesion');
        }
        //Se carga el helper html para usar en la vista
        $this->load->helper('html');
        //Se carga form_validation para validar los campos
        //$this->load->library('form_validation');
        //Se carga el modelo de los pagos en la base de datos
        $this->load->model('pago_model');
    }//Fin construct
    
    /**
     * Muestra la vista principal del m&oacute;dulo.
     * 
     * @access	public
     */
    function index(){
        //Se obtiene mediante la url el n&uacute;mero del contrato
        $this->data['id_contrato'] = $this->uri->segment(3);
        $id_contrato = $this->uri->segment(3);
        //Se traen los pagos existentes de ese contrato
        $this->data['pagos'] = $this->pago_model->listar_pagos($id_contrato); 
        //Se obtiene el modelo del estado de los pagos de ese contrato
        $this->data['estado_pagos'] = $this->pago_model->estado_pagos_contrato($id_contrato);
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Pagos';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'pago/pago_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
    
    /**
     * Funci&oacute;n que se encarga de realizar un pago al contrato
     * 
     * @access	public
     */
    function agregar(){
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules('fecha_pago', '', 'trim');
        $this->form_validation->set_rules('numero_acta', '', 'trim');
        $this->form_validation->set_rules('valor', 'El valor del pago', 'required|numeric|trim');
        $this->form_validation->set_rules('retenido', 'El valor retenido', 'numeric|trim');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
        $this->form_validation->set_message('numeric', '-%s no puede contener letras');
        
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
             * Se preparan los datos del contrato para enviar a la base de datos:
             * 
             */
            $pago = array(
                'Fecha' => $this->input->post('fecha_pago'),
                'Numero_Acta' => $this->input->post('numero_acta'),
                'Valor_Pago' => $this->input->post('valor'),
                'Valor_Retenido' => $this->input->post('retenido'),
                'Fk_Id_Contratos' => $this->input->post('id_contrato')
            );
            
            //Se obtienen los ids que se usar&aacute;n para la acci&oacute;n de auditor&iacute;a
            $id_contrato = $this->input->post('id_contrato');
            $valor_pago = $this->input->post('valor');
            $numero = $this->input->post('numero_contrato');
            
            //Se env&iacute;n los datos del pago al modelo
            $this->pago_model->agregar_pago($pago);
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->agregar_pago($valor_pago, $numero);
            
            //Se establece el mensaje de exito
            $this->data['mensaje_exito'] = 'Se ha realizado el pago por $'.number_format($valor_pago, 0, '', '.').' correctamente';
            
            /*
             * Se prepara la redirecci&oacute;n
             */
            //Se obtiene mediante la url el n&uacute;mero del contrato
            $this->data['id_contrato'] = $id_contrato;
            //Se traen los pagos existentes de ese contrato
            $this->data['pagos'] = $this->pago_model->listar_pagos($id_contrato); 
            //Se obtiene el modelo del estado de los pagos de ese contrato
            $this->data['estado_pagos'] = $this->pago_model->estado_pagos_contrato($id_contrato);
            //se establece el titulo de la p&aacute;gina
            $this->data['titulo'] = 'Contratos - Pagos';
            //se establece la vista que tiene el contenido principal
            $this->data['contenido_principal'] = 'pago/pago_view';
            //se carga el template
            $this->load->view('includes/template', $this->data);
        }
    }//Fin agregar
    
    /**
     * Funci&oacute;n que se encarga de eliminar un pago al contrato
     * 
     * @access	public
     */
    function eliminar(){
        //Se obtiene el id del contrato y el pago a eliminar
        $id_pago = $this->uri->segment(3);
        $id_contrato = $this->uri->segment(4);
        
        /*
         * Para la acci&oacute;n de auditor&iacute;a se lista el pago y se obtiene
         * el valor y el n&uacute;mero del contrato:
         */
        $pagos = $this->pago_model->listar_pagos($id_contrato);
        
        foreach ($pagos as $pago):
            $valor = number_format($pago->Valor_Pago, 0, '', '');
            $numero = $pago->Numero;
        endforeach;
        
        //Se env&iacute;n los datos del pago a eliminar al modelo
        $this->pago_model->eliminar_pago($id_pago);
       
        //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
        $this->auditoria_model->eliminar_pago($valor, $numero);
        
        //Se establece el mensaje de exito
        $this->data['mensaje_exito'] = 'Se ha eliminado el pago por '.$valor.' correctamente';
        
        /*
        * Se prepara la redirecci&oacute;n
        */
        //Se obtiene mediante la url el n&uacute;mero del contrato
        $this->data['id_contrato'] = $this->uri->segment(4);
        //Se traen los pagos existentes de ese contrato
        $this->data['pagos'] = $this->pago_model->listar_pagos($id_contrato); 
        //Se obtiene el modelo del estado de los pagos de ese contrato
        $this->data['estado_pagos'] = $this->pago_model->estado_pagos_contrato($id_contrato);
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Contratos - Pagos';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'pago/pago_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin eliminar()
}//Fin Pago
/* End of file pago.php */
/* Location: ./contratos/application/controllers/pago.php */
