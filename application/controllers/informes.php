<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

//Se carga la libreria de PDF
require('contratos/libraries/Fpdf.php');

//Se carga la librer&iacute;a de conversion de numeros a letras
require('contratos/libraries/Numeros.php');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo de informes.
 * @author 				John Arley Cano Salinas
 * @copyright           &copy; HATOVIAL S.A.S.
 */
Class Informes extends CI_Controller{
	/**
    * Funci&oacute;n constructora de la clase informes.
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
        //Se cargan la librer&iacute;a de PDF
        $this->load->library('Fpdf');
        //Se cargan la librer&iacute;a de Excel y la libreria que construye los graficos
        $this->load->library(array('PHPExcel', 'PDF_Diag', 'PHPWord'));
        //Se cargan los modelos
        $this->load->model(array('informes_model', 'contrato_model', 'email_model', 'pago_model', 'tercero_model'));
        //Se establece la ruta para las fuentes del reporte
        define('FPDF_FONTPATH','application/font/');
    }//Fin construct()

    /**
     * Muestra la vista principal del m&oacute;dulo informes.
     *
     * @access	public
     */
    function index(){
        //Se traen los estados de los contratos
        $this->data['contratos_estados'] = $this->contrato_model->listar_contratos_estados();
        //Se traen los contratistas que se encuentran en la aplicaci&oacute;n
        $this->data['contratistas'] = $this->tercero_model->listar_contratistas();
        //Se traen los contratantes
        $this->data['contratantes'] = $this->tercero_model->listar_contratantes();
        //se establece el titulo de la p&aacute;gina
        $this->data['titulo'] = 'Informes';
		// listado de contratos
		$this->data['contratos'] = $this->contrato_model->listar_contratos('');
        //se establece la vista que tiene el contenido principal
        $this->data['contenido_principal'] = 'informes/informes_view';
        //se carga el template
        $this->load->view('includes/template', $this->data);
    }//Fin index()

    /**
     * Acta de inicio en Word
     */
    function acta_inicio(){
        //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
        $this->auditoria_model->generar_acta_inicio($this->uri->segment(3));

        //Se carga la vista que contiene el reporte
        $this->load->view('informes/informes/acta_inicio');
    } // acta_inicio

    /**
     * Acta de recibo en Word
     */
    function acta_recibo(){
        //Se usa el modelo para la acci&oacute;n de auditor&iacute;a
        $this->auditoria_model->generar_acta_recibo($this->uri->segment(3));

        //Se carga la vista que contiene el reporte
        $this->load->view('informes/informes/acta_recibo');
    } // acta_recibo

    /**
     * Estructura del reporte de valores
     *
     * @access  public
     */
    function contratistas(){
        //Se recibe el usuario que viene por post
        $id_contratista = $this->input->post('id_contratista');

        //Nombre genérico
        $this->data['nombre'] = "TODOS LOS CONTRATISTAS";

        //Si viene id de contratista
        if ($id_contratista != "") {
            //Consulta de datos del contratista
            $contratista = $this->informes_model->obtener_nombre_contratista($id_contratista);

            //Nombre del contratista
            $this->data['nombre'] = $contratista->Nombre;
        };

        //Consulta los contratos asociados a ese contratista
        $this->data['contratos'] = $this->informes_model->listar_valores($id_contratista);

        //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/contratista', $this->data);
    }//Fin valores()

    /**
     * Estructura del reporte de contratantes en excel
     *
     * @access  public
     */
    function contratantes_excel(){
        //Se recibe el usuario que viene por post
        $id_contratante = $this->input->post('id_contratante');

        //Nombre genérico
        $this->data['nombre'] = "TODOS LOS CONTRATANTES";

        //Si viene id de contratantes
        if ($id_contratante != "") {
            //Consulta de datos del contratantes
            $contratantes = $this->informes_model->obtener_nombre_contratista($id_contratante);

            //Nombre del contratante
            $this->data['nombre'] = $contratantes->Nombre;
        };

        //Consulta los contratos asociados a ese contratantes
        $this->data['contratos'] = $this->informes_model->listar_valores("contratante", $id_contratante);

        // //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/contratista_excel', $this->data);
    }//Fin contratantes_excel()

    /**
     * Estructura del reporte de valores
     *
     * @access  public
     */
    function contratistas_excel(){
        //Se recibe el usuario que viene por post
        $id_contratista = $this->input->post('id_contratista');

        //Nombre genérico
        $this->data['nombre'] = "TODOS LOS CONTRATISTAS";

        //Si viene id de contratista
        if ($id_contratista != "") {
            //Consulta de datos del contratista
            $contratista = $this->informes_model->obtener_nombre_contratista($id_contratista);

            //Nombre del contratista
            $this->data['nombre'] = $contratista->Nombre;
        };

        //Consulta los contratos asociados a ese contratista
        $this->data['contratos'] = $this->informes_model->listar_valores("contratista", $id_contratista);

        //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/contratista_excel', $this->data);
    }//Fin contratistas_excel()

    /**
     * Estructura del reporte de valores
     *
     * @access	public
     */
    function valores(){
        //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/valores');
    }//Fin valores()

    /**
     * Estructura del acta de liquidacion
     *
     * @access  public
     */
    function acta_liquidacion(){
         //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/acta_liquidacion');
    }//Fin acta_liquidacion()

    /**
     * Estructura del acta de liquidacion en Word
     *
     * @access  public
     */
    function acta_liquidacion_word(){
         //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/acta_liquidacion_word');
    }//Fin acta_liquidacion_word()

    /**
     * Exporta a excel el listado de contratos seg&uacute;n el estado seleccionado.
     *
     * @access	public
     */
    function estados(){
        $this->form_validation->set_rules('estados_contratos', 'El estado del contrato', 'required|trim');
        $this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Es necesario que seleccione un estado<br/>para poder generar el informe";
            $this->index();
        }else{
             //Se carga la vista que contiene el informe
            $this->load->view('informes/informes/estados');
        }
    }//Fin estados()

    /**
     * Exporta a excel el listado de contratos seg&uacute;n el rango de fechas seleccionado.
     *
     * @access	public
     */
    function fecha_inicial(){
        //Se ejecutan las validaciones
        $this->form_validation->set_rules('fecha1', '', 'required|trim');
        $this->form_validation->set_rules('fecha2', '', 'required|trim');
        $this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Para poder filtrar por inicial final<br/>seleccione ambas fechas";
            $this->index();
        }else{
            //Se carga la vista que contiene el informe
            $this->load->view('informes/informes/fecha_inicio');
        }
    }//Fin fecha_inicial()

    /**
     * Exporta a excel el listado de contratos seg&uacute;n el rango de fechas seleccionado.
     *
     * @access	public
     */
    function fecha_vencimiento(){
        //Se ejecutan las validaciones
        $this->form_validation->set_rules('fecha3', '', 'required|trim');
        $this->form_validation->set_rules('fecha4', '', 'required|trim');
        $this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
        if($this->form_validation->run() == false){
            //Se imprime el mensaje de informaci&oacute;n
            $this->data['mensaje_alerta'] = "Para poder filtrar por fecha de vencimiento<br/>seleccione ambas fechas";
            $this->index();
        }else{
            //Se carga la vista que contiene el informe
            $this->load->view('informes/informes/fecha_vencimiento');
        }
    }//Fin fecha_vencimiento()

    /**
     * Exporta a excel el listado de contratos seg&uacute;n el rango de fechas seleccionado.
     *
     * @access	public
     */
    function no_acta_inicio(){
        //Se carga la vista que contiene el informe
        $this->load->view('informes/informes/no_acta_inicio');
    }//Fin no_acta_inicio

	function pagos() {
		if ($this->uri->segment(3) == '') {
			$this->data['id_contrato'] = $this->input->post('id_contrato');
			$this->form_validation->set_rules('id_contrato', 'Contrato', 'required|trim');
			$this->form_validation->set_message('required', '-%s no puede estar vac&iacute;o');
			if($this->form_validation->run() == false){
				//Se imprime el mensaje de informaci&oacute;n
				$this->data['mensaje_alerta'] = "Es necesario que seleccione un contrato<br/>para poder generar el informe";
				$this->index();
				return 0;
			}
		} else {
			$this->data['id_contrato'] = $this->uri->segment(3);
		}
		$id_contrato = $this->data['id_contrato'];
		//Se traen los pagos existentes de ese contrato
		$this->data['pagos'] = $this->pago_model->listar_pagos($id_contrato);
		//Se obtiene el modelo del estado de los pagos de ese contrato
		$this->data['estado_pagos'] = $this->pago_model->estado_pagos_contrato($id_contrato);
		$this->load->view('informes/informes/pagos', $this->data);
	}
}//Fin informes

/**
 * Clase encargada de armar las cabeceras y los pi&eacute; de p&aacute;gina de
 * los informes en PDF.
 * @author 		John Arley Cano Salinas
 * @copyright           &copy; HATOVIAL S.A.S.
 */
class PDF extends FPDF {

    /**
     * Construye la cabecera para los documentos en PDF
     *
     * @access	public
     */
    function Header(){
        // Logo
        //$this->Image('logo_pb.png',10,8,33);
        //
        // Arial bold 15
        $this->SetFont('Arial','B',15);

        // Movernos a la derecha
        //$this->Cell(0);

        // Título
        $this->MultiCell(0,7, utf8_decode('ACTA DE LIQUIDACIÓN DE OBRA'),1,'C');
        //$pdf->MultiCell(0,6, utf8_decode('OFERTA MERCANTIL NRO. '.$contrato->Numero),0,'C');

        // Salto de línea
        $this->Ln(10);
    }//Fin header()

    /**
     * Construye el pi&eacute; de p&aacute;gina para los documentos en PDF
     *
     * @access	public
     */
    function Footer(){
        //Posici&oacute;n: a 1,5 cm del final
        $this->SetY(-15);

        //Tipo de letra
        $this->SetFont('Arial','',8);

        // Número de p&aacute;gina
        $this->Cell(0,10, utf8_decode('Sistema de Gestión de Contratos - Hatovial S.A.S. Página ').$this->PageNo().' de {nb}',0,0,'R');
    }//Fin footer()
}//Fin PDF
/* End of file informes.php */
/* Location: ./contratos/application/controllers/Informes.php */
