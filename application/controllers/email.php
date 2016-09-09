<?php
//Zona horaria
date_default_timezone_set('America/Bogota');

if ( ! defined('BASEPATH')) exit('Lo sentimos, usted no tiene acceso a esta ruta');

/**
 * Controlador del m&oacute;dulo que env&iacute;a  correos electr&oacute;nicos.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Email extends CI_Controller{
    /**
    * Funci&oacute;n constructora de la clase email. 
    * 
    * Se hereda el mismo constructor de la clase Controller para evitar sobreescribirlo y de esa manera 
    * conservar el funcionamiento de controlador.
    * 
    * @access	public
    */
    function __construct() {
        //con esta linea se hereda el constructor de la clase Controller
        parent::__construct();
        //Se cargan los modelos
        $this->load->model('email_model');
        $this->load->model('contrato_model');
        $this->load->model('pago_model');
    }//Fin construct()
    
    /**
     * Ejecuta los m&eacute;todos que env&iacute;n los correos.
     * 
     * @access	public
     */
    function index(){
        /*
         * Se llaman todos los m&eacute;todos para que se invoque autom&aacute;ticamente
         * todos los correos desde solo un lugar
         */
        $this->_contratos_en_vencimiento();
        $this->_contratos_vencidos();
        $this->polizas_en_vencimiento();
        $this->polizas_vencidas();
        $this->pagos_excedidos();
        $this->no_acta_inicio();
    }//Fin index
    
    /**
     * Verifica los contratos que est&aacute;n a punto de vencerse.
     * 
     * @access	private
     */
    function _contratos_en_vencimiento(){
        //Se carga el modelo que lista los contratos en vencimiento
        $contratos = $this->email_model->contratos_en_vencimiento();
        
        /*
         * Se construye la tabla que se va a enviar
         */
        /********************************tabla********************************/
        $tabla = '<table border="1" bordercolor="white" cellspacing="0" width="100%">';
        /********************************cabecera********************************/
        $tabla .= '<thead border="1" bordercolor="white" style="background-color:#0B37B0; font-family:Tahoma; color: white; font-size: 13px;" width="100%">';
        $tabla .= '<tr>';
        $tabla .= '<th>Contrato</th>';
        $tabla .= '<th>Objeto</th>';
        $tabla .= '<th>Contratista</th>';
        $tabla .= '<th>Fecha de vencimiento</th>';
        $tabla .= '<th>D&iacute;as restantes</th>';
        $tabla .= '</tr>';
        $tabla .= '</thead>';
        /********************************cuerpo********************************/
        $tabla .= '<tbody style="font-family:Tahoma; font-size: 12px;">';
            $contador = 0;
            foreach($contratos as $contrato):
                $contador++;
                $tabla .= '<tr bordercolor="black">';
                $tabla .= '<td align="right">'.$contrato->Numero.'&nbsp;</td>';
                $tabla .= '<td>'.$contrato->Objeto.'</td>';
                $tabla .= '<td>'.$contrato->Contratista.'</td>';
                $tabla .= '<td width="12%">'.$contrato->Fecha_Vencimiento.'</td>';
                $tabla .= '<td align="right">'.$contrato->Dias_Restantes.'</td>';
                $tabla .= '</tr>';
            endforeach;
        $tabla .= '</tbody>';
        $tabla .= '</table>';
        
        //Se define el asunto
        $asunto = 'Contratos por vencerse';
        
        //Se verifica, si hay datos se env&iacute;a la tabla
        if($contador == 0){
            $cuerpo = 
            '<b>Listado de contratos por vencerse</b><br/><br/>A la fecha no hay contratos por vencerse.';
        }else{
            $cuerpo = '<b>Listado de contratos por vencerse</b><br/><br/>Este es el listado de los '.$contador.' contratos que est&aacute;n por vencerse dentro de los pr&oacute;ximos cinco d&iacute;as:'.$tabla;
        }
        
        //Se ejecuta el modelo que env&iaacute;
        $this->email_model->_enviar_email($asunto, $cuerpo);
        
        //Mensaje de &eacute;xito
        echo 'El mensaje de contratos por vencerse se ha enviado correctamente<br/>';
    }//Fin _contratos_en_vencimiento()
    
    /**
     * Verifica los contratos que est&aacute;n vencidos.
     * 
     * @access	private
     */
    function _contratos_vencidos(){
        //Se carga el modelo que trae los contratos vencidos
        $contratos = $this->email_model->contratos_vencidos();
        
        /*
         * Se construye la tabla que se va a enviar
         */
        /********************************tabla********************************/
        $tabla = '<table border="1" bordercolor="white" cellspacing="0" width="100%">';
        /********************************cabecera********************************/
        $tabla .= '<thead border="1" bordercolor="white" style="background-color:#0B37B0; font-family:Tahoma; color: white; font-size: 13px;" width="100%">';
        $tabla .= '<tr>';
        $tabla .= '<th>Contrato</th>';
        $tabla .= '<th>Objeto</th>';
        $tabla .= '<th>Contratista</th>';
        $tabla .= '<th>Fecha de vencimiento</th>';
        $tabla .= '<th>D&iacute;as vencido</th>';
        $tabla .= '</tr>';
        $tabla .= '</thead>';
        /********************************cuerpo********************************/
        $tabla .= '<tbody style="font-family:Tahoma; font-size: 12px;">';
            $contador = 0;
            foreach($contratos as $contrato):
                $contador++;
                $tabla .= '<tr bordercolor="black">';
                $tabla .= '<td align="right">'.$contrato->Numero.'&nbsp;</td>';
                $tabla .= '<td>'.$contrato->Objeto.'</td>';
                $tabla .= '<td>'.$contrato->Contratista.'</td>';
                $tabla .= '<td width="12%">'.$contrato->Fecha_Vencimiento.'</td>';
                $tabla .= '<td align="right">'.$contrato->Dias_Vencidos.'</td>';
                $tabla .= '</tr>';
            endforeach;
        $tabla .= '</tbody>';
        $tabla .= '</table>';
        
        //Se define el asunto
        $asunto = 'Contratos vencidos';
        
        //Se verifica, si hay datos se env&iacute;a la tabla
        if($contador == 0){
            $cuerpo = 
            '<b>Listado de contratos vencidos</b><br/><br/>A la fecha no hay contratos vencidos.';
        }else{
            $cuerpo = '<b>Listado de contratos vencidos</b><br/><br/>Este es el listado de los '.$contador.' contratos que est&aacute;n vencidos a la fecha:'.$tabla;
        }
        
        //Se ejecuta el modelo que env&iaacute;
        $this->email_model->_enviar_email($asunto, $cuerpo);
        
        //Mensaje de &eacute;xito
        echo 'El mensaje de contratos vencidos se ha enviado correctamente<br/>';
    }//Fin _contratos_vencidos()
    
    /**
     * Verifica las p&oacute;lizas que est&eacute;n a punto de vencerse.
     * 
     * @access	private
     */
    function polizas_en_vencimiento(){
        //Se carga el modelo que se utilizar&aacute;
        $polizas = $this->email_model->polizas_en_vencimiento();
        
        /*
        * Se construye la tabla que se va a enviar
        */
        /********************************tabla********************************/
        $tabla = '<table border="1" bordercolor="white" cellspacing="0" width="100%">';
        /********************************cabecera********************************/
        $tabla .= '<thead border="1" bordercolor="white" style="background-color:#0B37B0; font-family:Tahoma; color: white; font-size: 13px;" width="100%">';
        $tabla .= '<tr>';
        $tabla .= '<th>Contrato</th>';
        $tabla .= '<th>P&oacute;liza</th>';
        $tabla .= '<th>Objeto</th>';
        $tabla .= '<th>Contratista</th>';
        $tabla .= '<th>Fecha de vencimiento</th>';
        $tabla .= '<th>D&iacute;as restantes</th>';
        $tabla .= '</tr>';
        $tabla .= '</thead>';
        /********************************cuerpo********************************/
        $tabla .= '<tbody style="font-family:Tahoma; font-size: 12px;">';
            $contador = 0;
            foreach ($polizas as $poliza):
                $contador++;
                $tabla .= '<tr bordercolor="black">';
                $tabla .= '<td>'.$poliza->Numero.'</td>';
                $tabla .= '<td>'.$poliza->Poliza_Tipo.'</td>';
                $tabla .= '<td>'.$poliza->Objeto.'</td>';
                $tabla .= '<td>'.$poliza->Contratista.'</td>';
                $tabla .= '<td width="12%">'.$poliza->Fecha_Final.'</td>';
                $tabla .= '<td align="right">'.$poliza->Dias_Vencidos.'</td>';
                $tabla .= '</tr>';
            endforeach;
        $tabla .= '</tbody>';
        $tabla .= '</table>';

        //Se define el asunto
        $asunto = 'Pólizas por vencerse';

        //Se verifica, si hay datos se env&iacute;a la tabla
        if($contador == 0){
            $cuerpo = 
            '<b>Listado de p&oacute;lizas por vencerse</b><br/><br/>A la fecha no hay p&oacute;lizas por vencerse.';
        }else{
            $cuerpo = '<b>Listado de p&oacute;lizas por vencerse</b><br/><br/>Este es el listado de las '.$contador.' p&oacute;lizas que est&aacute;n por vencerse dentro de los pr&oacute;ximos cinco d&iacute;as:'.$tabla;
        }

        //Se ejecuta el modelo que env&iaacute;
        $this->email_model->_enviar_email($asunto, $cuerpo);
        
        //Mensaje de &eacute;xito
        echo 'El mensaje de p&oacute;lizas por vencerse se ha enviado correctamente<br/>';
    }//Fin polizas_en_vencimiento()
    
    /**
     * Verifica las p&oacute;lizas que est&eacute;n vencidas.
     * 
     * @access	private
     */
    function polizas_vencidas(){
        //Se carga el modelo que se utilizar&aacute;
        $polizas = $this->email_model->polizas_vencidas();
        
        /*
        * Se construye la tabla que se va a enviar
        */
        /********************************tabla********************************/
        $tabla = '<table border="1" bordercolor="white" cellspacing="0" width="100%">';
        /********************************cabecera********************************/
        $tabla .= '<thead border="1" bordercolor="white" style="background-color:#0B37B0; font-family:Tahoma; color: white; font-size: 13px;" width="100%">';
        $tabla .= '<tr>';
        $tabla .= '<th>Contrato</th>';
        $tabla .= '<th>P&oacute;liza</th>';
        $tabla .= '<th>Objeto</th>';
        $tabla .= '<th>Contratista</th>';
        $tabla .= '<th>Fecha de vencimiento</th>';
        $tabla .= '<th>D&iacute;as vencidos</th>';
        $tabla .= '</tr>';
        $tabla .= '</thead>';
        /********************************cuerpo********************************/
        $tabla .= '<tbody style="font-family:Tahoma; font-size: 12px;">';
            $contador = 0;
            foreach ($polizas as $poliza):
                $contador++;
                $tabla .= '<tr bordercolor="black">';
                $tabla .= '<td>'.$poliza->Numero.'</td>';
                $tabla .= '<td>'.$poliza->Poliza_Tipo.'</td>';
                $tabla .= '<td>'.$poliza->Objeto.'</td>';
                $tabla .= '<td>'.$poliza->Contratista.'</td>';
                $tabla .= '<td width="12%">'.$poliza->Fecha_Final.'</td>';
                $tabla .= '<td align="right">'.$poliza->Dias_Vencidos.'</td>';
                $tabla .= '</tr>';
            endforeach;
        $tabla .= '</tbody>';
        $tabla .= '</table>';

        //Se define el asunto
        $asunto = 'Pólizas vencidas';
        
        //Se verifica, si hay datos se env&iacute;a la tabla
        if($contador == 0){
            $cuerpo = 
            '<b>Listado de p&oacute;lizas vencidas</b><br/><br/>A la fecha no hay p&oacute;lizas vencidas.';
        }else{
            $cuerpo = '<b>Listado de contratos vencidos</b><br/><br/>Este es el listado de las '.$contador.' p&oacute;lizas que est&aacute;n vencidas a la fecha:'.$tabla;
        }

        //Se ejecuta el modelo que env&iaacute;
        $this->email_model->_enviar_email($asunto, $cuerpo);
        
        //Mensaje de &eacute;xito
        echo 'El mensaje de p&oacute;lizas vencidas se ha enviado correctamente<br/>';
    }//Fin polizas_vencidas()
    
    /**
     * Verifica los contratos que tengan pagos mayores al valor del contrato.
     * 
     * @access	private
     */
    function pagos_excedidos(){
        //Se carga el modelo
        $pagos = $this->pago_model->pagos_excedidos();
        
        /*
        * Se construye la tabla que se va a enviar
        */
        /********************************tabla********************************/
        $tabla = '<table border="1" bordercolor="white" cellspacing="0" width="100%">';
        /********************************cabecera********************************/
        $tabla .= '<thead border="1" bordercolor="white" style="background-color:#0B37B0; font-family:Tahoma; color: white; font-size: 13px;" width="100%">';
        $tabla .= '<tr>';
        $tabla .= '<th>Contrato</th>';
        $tabla .= '<th>Objeto</th>';
        $tabla .= '<th>Contratista</th>';
        $tabla .= '<th>Valor del contrato</th>';
        $tabla .= '<th>Valor pagado</th>';
        $tabla .= '<th>Valor excedido</th>';
        $tabla .= '</tr>';
        $tabla .= '</thead>';
        /********************************cuerpo********************************/
        $tabla .= '<tbody style="font-family:Tahoma; font-size: 12px;">';
            $contador = 0;
            foreach ($pagos as $pago):
                if($pago->Excedido > 0){
                   $contador++;
                    $tabla .= '<tr bordercolor="black">';
                    $tabla .= '<td>'.$pago->Numero.'</td>';
                    $tabla .= '<td>'.$pago->Objeto.'</td>';
                    $tabla .= '<td>'.$pago->Contratista.'</td>';
                    $tabla .= '<td align="right" width="15%">'.'$ '.number_format($pago->Valor_Total, 0, '', '.').'</td>';
                    $tabla .= '<td align="right" width="15%">'.'$ '.number_format($pago->Pagado, 0, '', '.').'</td>';
                    $tabla .= '<td align="right" width="15%"><b>'.'$ '.number_format($pago->Excedido, 0, '', '.').'</b></td>';
                    $tabla .= '</tr>'; 
                }
            endforeach;
        $tabla .= '</tbody>';
        $tabla .= '</table>';
        
        //Se define el asunto
        $asunto = 'Pagos excedidos a contratos';
        
        //Se verifica, si hay datos se env&iacute;a la tabla
        if($contador == 0){
            $cuerpo = 
            '<b>Listado de pagos excedidos a los contratos</b><br/><br/>A la fecha no hay ning&uacute;n pago que exceda el valor de un contrato.';
        }else{
            $cuerpo = '<b>Listado de pagos excedidos a los contratos</b><br/><br/>Este es el listado de los '.$contador.' contratos que tienen pagos que van por encima de su valor y aun no han sido liquidados:'.$tabla;
        }

        //Se ejecuta el modelo que env&iaacute;
        $this->email_model->_enviar_email($asunto, $cuerpo);
        
        //Mensaje de &eacute;xito
        echo 'El mensaje de pagos excedidos se ha enviado correctamente<br/>';
    }//Fin pagos_excedidos
    
    /**
     * Verifica los contratos que No tienen acta de inicio
     * 
     * @access	public
     */
    function no_acta_inicio(){
        //Se carga el modelo
        $contratos = $this->email_model->no_acta_inicio();
        
        /*
        * Se construye la tabla que se va a enviar
        */
        /********************************tabla********************************/
        $tabla = '<table border="1" bordercolor="white" cellspacing="0" width="100%">';
        /********************************cabecera********************************/
        $tabla .= '<thead border="1" bordercolor="white" style="background-color:#0B37B0; font-family:Tahoma; color: white; font-size: 13px;" width="100%">';
        $tabla .= '<tr>';
        $tabla .= '<th>Contrato</th>';
        $tabla .= '<th>Objeto</th>';
        $tabla .= '<th>Contratista</th>';
        $tabla .= '<th>Fecha inicial</th>';
        $tabla .= '<th>Fecha vencimiento</th>';
        $tabla .= '<th>Valor inicial</th>';
        $tabla .= '</tr>';
        $tabla .= '</thead>';
        /********************************cuerpo********************************/
        $tabla .= '<tbody style="font-family:Tahoma; font-size: 12px;">';
            $contador = 0;
            foreach ($contratos as $contrato):
                $contador++;
                $tabla .= '<tr bordercolor="black">';
                $tabla .= '<td align="right">'.$contrato->Numero.'</td>';
                $tabla .= '<td>'.$contrato->Objeto.'</td>';
                $tabla .= '<td>'.$contrato->Contratista.'</td>';
                $tabla .= '<td align="right" width="15%">'.$contrato->Fecha_Inicial.'</td>';
                $tabla .= '<td align="right" width="15%">'.$contrato->Fecha_Vencimiento.'</td>';
                $tabla .= '<td align="right" width="15%">$ '.number_format($contrato->Valor_Inicial, 0, '', '.').'</td>';
                $tabla .= '</tr>';
            endforeach;
        $tabla .= '</tbody>';
        $tabla .= '</table>';
        
        //Se define el asunto
        $asunto = 'Contratos sin acta de inicio';
        
        //Se verifica, si hay datos se env&iacute;a la tabla
        if($contador == 0){
            $cuerpo = 
            '<b>Listado de contratos que no tienen acta de inicio</b><br/><br/>A la fecha todos los contratos poseen acta de inicio.';
        }else{
            $cuerpo = '<b>Listado de contratos que no tienen acta de inicio</b><br/><br/>Este es el listado de los '.$contador.' contratos que no tienen acta de inicio:'.$tabla;
        }
        
        //Se ejecuta el modelo que env&iaacute;
        $this->email_model->_enviar_email($asunto, $cuerpo);
        
        //Mensaje de &eacute;xito
        echo 'El mensaje de contratos sin acta de inicio se ha enviado correctamente<br/>';
    }//Fin no_acta_inicio
}//Fin email
/* End of file email.php */
/* Location: ./contratos/application/controllers/email.php */