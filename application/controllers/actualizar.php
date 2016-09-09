<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo Modificar Contrato.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Actualizar extends CI_Controller{
    
    /**
    * Funci&oacute;n constructora de la clase actualizar. 
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
        //Se carga el modelo de los contratistas en la base de datos
        $this->load->model('tercero_model');
        //Se carga el modelo de los contratos en la base de datos
        $this->load->model('contrato_model');
        //Se carga el helper html para usar en la vista
        $this->load->helper('html');
    }//Fin construct()
    
    /**
     * Muestra la vista principal del m&oacute;dulo Actualizar.
     * 
     * @access	public
     */
    function index(){
        //Se obtiene mediante la url el n&uacute;mero del contrato
        $id_contrato = $this->uri->segment(3);
        //Se traen todos los contratos
        $this->data['contratos'] = $this->contrato_model->listar_contratos($id_contrato);
        //Se traen los contratantes
        $this->data['contratantes'] = $this->tercero_model->listar_contratantes();
        //Se traen los contratistas que se encuentran en la aplicaci&oacute;n
        $this->data['contratistas'] = $this->tercero_model->listar_contratistas();
        //Se traen los estados de los contratos
        $this->data['contratos_estados'] = $this->contrato_model->listar_contratos_estados();
        //Se traen los datos de la poliza de cumplimiento
        $this->data['poliza_cumplimiento'] = $this->contrato_model->listar_poliza_cumplimiento($id_contrato);
        //Se traen los datos de la poliza de prestaciones sociales
        $this->data['poliza_prestaciones'] = $this->contrato_model->listar_poliza_prestaciones($id_contrato);
        //Se traen los datos de la poliza de anticipos
        $this->data['poliza_anticipos'] = $this->contrato_model->listar_poliza_anticipos($id_contrato);
        //Se traen los datos de la poliza de calidad
        $this->data['poliza_calidad'] = $this->contrato_model->listar_poliza_calidad($id_contrato);
        //Se traen los datos de la poliza de estabilidad
        $this->data['poliza_estabilidad'] = $this->contrato_model->listar_poliza_estabilidad($id_contrato);
        //Se traen los datos de la poliza de responsabilidad civil contractual
        $this->data['poliza_rc'] = $this->contrato_model->listar_poliza_rc($id_contrato);
        //Se traen las aseguradoras de los contratos
        $this->data['aseguradoras'] = $this->tercero_model->listar_aseguradoras();
        //Se traen las adiciones pertenecientes a ese contrato
        $this->data['adiciones'] = $this->contrato_model->listar_adiciones($id_contrato);
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Actualizar contrato';
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'actualizar/actualizar_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()
    
    /**
     * Funci&oacute;n que se encarga de actualizar el contrato
     * 
     * @access	public
     */
    function contrato(){
        /*
         * Este array se usa para mantener los datos modificados de todos los campos
         * en caso de no superar la validaci&oacute;n
         */
        $datos = array(
            array('field' => 'fecha_inicial', 'label' => '', 'rules' => 'trim'),
            array('field' => 'fecha_acta', 'label' => '', 'rules' => 'trim'),
            array('field' => 'acta_inicio', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_cumplimiento', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_prestaciones', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_anticipos', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_calidad', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_estabilidad', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_estabilidad', 'label' => '', 'rules' => 'trim'),
            array('field' => 'numero_rc', 'label' => '', 'rules' => 'trim'),
            array('field' => 'inicio_cumplimiento', 'label' => '', 'rules' => 'trim'),
            array('field' => 'inicio_prestaciones', 'label' => '', 'rules' => 'trim'),
            array('field' => 'inicio_anticipos', 'label' => '', 'rules' => 'trim'),
            array('field' => 'inicio_calidad', 'label' => '', 'rules' => 'trim'),
            array('field' => 'inicio_rc', 'label' => '', 'rules' => 'trim'),
            array('field' => 'vencimiento_cumplimiento', 'label' => '', 'rules' => 'trim'),
            array('field' => 'vencimiento_prestaciones', 'label' => '', 'rules' => 'trim'),
            array('field' => 'vencimiento_anticipos', 'label' => '', 'rules' => 'trim'),
            array('field' => 'vencimiento_calidad', 'label' => '', 'rules' => 'trim'),
            array('field' => 'vencimiento_rc', 'label' => '', 'rules' => 'trim'),
            array('field' => 'aseguradora_cumplimiento', 'label' => '', 'rules' => 'trim'),
            array('field' => 'aseguradora_prestaciones', 'label' => '', 'rules' => 'trim'),
            array('field' => 'aseguradora_anticipos', 'label' => '', 'rules' => 'trim'),
            array('field' => 'aseguradora_calidad', 'label' => '', 'rules' => 'trim'),
            array('field' => 'aseguradora_rc', 'label' => '', 'rules' => 'trim'),
        );
        /*
         * Estas son las validaciones que se realizan con el Helper form_validation
         */
        $this->form_validation->set_rules($datos);
        if($this->input->post('nuevo_contratista') == ''){
            $this->form_validation->set_rules('contratista', 'El contratista', 'required|trim');
        }else{
            $this->form_validation->set_rules('nuevo_contratista', 'El contratista', 'trim');
        }
        $this->form_validation->set_rules('objeto_contrato', 'El objeto del contrato', 'required|trim');
        $this->form_validation->set_rules('localizacion_contrato', 'La localizaci&oacute;n del contrato', 'required|trim');
        $this->form_validation->set_rules('estados_contratos', 'El estado del contrato', 'required|trim');
        $this->form_validation->set_rules('valor_inicial', 'El valor inicial del contrato', 'numeric|trim');
        $this->form_validation->set_rules('plazo', 'El plazo', 'numeric|trim');
        $this->form_validation->set_rules('valor_cumplimiento', 'El valor de la p&oacute;liza de cumplimiento', 'numeric|trim');
        $this->form_validation->set_rules('valor_prestaciones', 'El valor de la p&oacute;liza de prestaciones', 'numeric|trim');
        $this->form_validation->set_rules('valor_anticipos', 'El valor de la p&oacute;liza de anticipos', 'numeric|trim');
        $this->form_validation->set_rules('valor_calidad', 'El valor de la p&oacute;liza de calidad', 'numeric|trim');
        $this->form_validation->set_rules('valor_estabilidad', 'El valor de la p&oacute;liza de estabilidad', 'numeric|trim');
        $this->form_validation->set_rules('valor_rc', 'El valor de la p&oacute;liza de responsabilidad civil contractual', 'numeric|trim');
        $this->form_validation->set_rules('plazo_adicion', 'El plazo de la adici&oacute;n', 'numeric|trim');
        $this->form_validation->set_rules('valor_adicion', 'El valor de la adici&oacute;n', 'numeric|trim');
        
        //Mensajes que se muestran cuando no se supera la validaci&oacute;n
        $this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
        $this->form_validation->set_message('numeric', '-%s no puede contener letras.');
        
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
            if($this->input->post('nuevo_contratista') != ''){
                $contratista = array(
                    'Fk_Id_Terceros_Tipo' => 2,
                    'Nombre' => $this->input->post('nuevo_contratista')
                );
                //Se envian los datos del nuevo contratista al modelo
                $this->tercero_model->registrar_tercero($contratista);
                
                //Se obtiene el id que se usar&aacute; como el Id del tercero
                //Y el nombre del contratista
                $id_contratista = $this->db->insert_id();
                $nombre = $this->input->post('nuevo_contratista');
                
                //Se envian los datos del contratista al modelo
                $this->auditoria_model->insertar_tercero($id_contratista, $nombre);
            }else{
                $id_contratista = $this->input->post('contratista');
            }

            //Si existe acta de inicio
            if($this->input->post('acta_inicio') == 1){
                //Esta sera la fecha de inicio del contrato
                $fecha_inicial = $this->input->post('fecha_acta');
            }else{
                $fecha_inicial = $this->input->post('fecha_inicial');
            }

            $plazo_total =  $this->input->post('plazo_adiciones') + $this->input->post('plazo') + $this->input->post('plazo_adicion');
            //Se ejecuta el modelo que calcula la fecha de vencmiento a partir de la fecha inicial y el plazo
            $fecha_vencimiento = $this->contrato_model->calcular_vencimiento($fecha_inicial, $plazo_total);
            
            $contrato = array(
                'Numero' => $this->input->post('numero_contrato'),
                'Objeto' => $this->input->post('objeto_contrato'),
                'Lugar' => $this->input->post('localizacion_contrato'),
                'Fk_Id_Estado' => $this->input->post('estados_contratos'),
                'Valor_Inicial' => $this->input->post('valor_inicial'),
                'Fecha_Inicial' => $this->input->post('fecha_inicial'),
                'Plazo' => $this->input->post('plazo'),
                'Fecha_Vencimiento' => $fecha_vencimiento,
                'Porcentaje_Avance' => $this->input->post('porcentaje_avance'),
                'Acta_Inicio' => $this->input->post('acta_inicio'),
                'Fecha_Acta_Inicio' => $this->input->post('fecha_acta'),
                'Fk_id_Usuario' => $this->session->userdata('Pk_Id_Usuario'),
                'Fk_Id_Terceros_Contratante' => $this->input->post('contratante'),
                'Fk_Id_Terceros' => $id_contratista
            );
            
            //Condici&oacute;n para encontrar el predio espec&iacute;fico
            $id_contrato = $this->input->post('id_contrato');
            $numero = $this->input->post('numero_contrato');
            
            /*
             * Se preparan los datos de las polizas del contrato
             * para enviar a la base de datos.
             * 
             */
            $poliza_cumplimiento = array(
                'Numero' => $this->input->post('numero_cumplimiento'),
                'Fecha_Inicio' => $this->input->post('inicio_cumplimiento'),
                'Fecha_Final' => $this->input->post('vencimiento_cumplimiento'),
                'Valor' => $this->input->post('valor_cumplimiento'),
                'Fk_Id_Tercero' => $this->input->post('aseguradora_cumplimiento'),
            );
            
            $poliza_prestaciones = array(
                'Numero' => $this->input->post('numero_prestaciones'),
                'Fecha_Inicio' => $this->input->post('inicio_prestaciones'),
                'Fecha_Final' => $this->input->post('vencimiento_prestaciones'),
                'Valor' => $this->input->post('valor_prestaciones'),
                'Fk_Id_Tercero' => $this->input->post('aseguradora_prestaciones'),
            );
            
            $poliza_anticipos = array(
                'Numero' => $this->input->post('numero_anticipos'),
                'Fecha_Inicio' => $this->input->post('inicio_anticipos'),
                'Fecha_Final' => $this->input->post('vencimiento_anticipos'),
                'Valor' => $this->input->post('valor_anticipos'),
                'Fk_Id_Tercero' => $this->input->post('aseguradora_anticipos'),
            );
            
            $poliza_calidad = array(
                'Numero' => $this->input->post('numero_calidad'),
                'Fecha_Inicio' => $this->input->post('inicio_calidad'),
                'Fecha_Final' => $this->input->post('vencimiento_calidad'),
                'Valor' => $this->input->post('valor_calidad'),
                'Fk_Id_Tercero' => $this->input->post('aseguradora_calidad'),
            );
            
            $poliza_estabilidad = array(
                'Numero' => $this->input->post('numero_estabilidad'),
                'Fecha_Inicio' => $this->input->post('inicio_estabilidad'),
                'Fecha_Final' => $this->input->post('vencimiento_estabilidad'),
                'Valor' => $this->input->post('valor_estabilidad'),
                'Fk_Id_Tercero' => $this->input->post('aseguradora_estabilidad'),
            );
            
            $poliza_rc = array(
                'Numero' => $this->input->post('numero_rc'),
                'Fecha_Inicio' => $this->input->post('inicio_rc'),
                'Fecha_Final' => $this->input->post('vencimiento_rc'),
                'Valor' => $this->input->post('valor_rc'),
                'Fk_Id_Tercero' => $this->input->post('aseguradora_rc'),
            );
            
            //Se toma los id de las p&oacute;lizas
            $id_cumplimiento = $this->input->post('id_cumplimiento');
            $id_prestaciones = $this->input->post('id_prestaciones');
            $id_anticipos = $this->input->post('id_anticipos');
            $id_calidad = $this->input->post('id_calidad');
            $id_estabilidad = $this->input->post('id_estabilidad');
            $id_rc = $this->input->post('id_rc');
            
            //Estos datos se env&iacute;an al contrato para modificarlo
            $this->contrato_model->actualizar_contrato($contrato, $id_contrato);
            
            //Se envian los datos de las polizas al modelo
            $this->contrato_model->actualizar_polizas(
                $poliza_cumplimiento, $id_cumplimiento, 
                $poliza_prestaciones, $id_prestaciones, 
                $poliza_anticipos, $id_anticipos, 
                $poliza_calidad, $id_calidad,
                $poliza_estabilidad, $id_estabilidad,
                $poliza_rc, $id_rc
            );

            //Se toman los datos para agregar la adicion
            if ($this->input->post('valor_adicion') || $this->input->post('plazo_adicion')) {
                $adicion = array(
                    'Valor' => $this->input->post('valor_adicion'),
                    'Plazo' => $this->input->post('plazo_adicion'),
                    'Fk_Id_Contrato' => $id_contrato
                );

                //Se inserta en la base de datos la nueva adicion
                $this->contrato_model->agregar_adicion($adicion);
            }//Fin if
            
            //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
            $this->auditoria_model->modificar_contrato($numero);
         
            //Se establece el mensaje de exito
            $this->data['mensaje_exito'] = 'El contrato '.$numero.' ha sido actualizado con &eacute;xito';
            
            /*
             * Se prepara la redirecci&oacute;n
             */
            $id_contrato = '';
            //Se traen todos los contratos
            $this->data['contratos'] = $this->contrato_model->listar_contratos($id_contrato);
            //Se traen los estados de los contratos
            $this->data['contratos_estados'] = $this->contrato_model->listar_contratos_estados();
            //se establece el titulo de la p&aacute;gina donde ir&aacute; despu&eacute;s de guardar
            $this->data['titulo'] = 'Contratos - Inicio';
            //se establece la vista que tiene el contenido principal
            $this->data['contenido_principal'] = 'inicio/inicio_view';
            //se carga el template
            $this->load->view('includes/template', $this->data);
        }
    }//Fin contrato()
}//Fin actualiazar
/* End of file actualizar.php */
/* Location: ./contratos/application/controllers/actualizar.php */