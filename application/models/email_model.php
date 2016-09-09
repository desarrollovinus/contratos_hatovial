<?php
/**
 * Modelo que se encarga de enviar los correos electr&oacute;nicos
 * @author 		John Arley Cano Salinas
 * @copyright	&copy;  HATOVIAL S.A.S.
 */
Class Email_model extends CI_Model{
    /*
     * Variables globales de configuraci&oacute;n del correo
     */
    var $protocolo = 'smtp';
    var $servidor_correo = 'mail.hatovial.com';
    var $usuario_sistema = 'contratos';
    var $password_sistema = 'C0ntrat0s';
    var $usuarios_correo = array('john.cano@hatovial.com');
    //var $usuarios_correo = array('johnarleycano@hotmail.com', 'oscar.alarcon@hatovial.com');
    var $correo_sistema = 'contratos@hatovial.com';
    
    /**
    * Env&iacute;a los correos electr&oacute;nicos
    * 
    * @access	private
    */
    function _enviar_email($asunto, $cuerpo){
        $this->load->library('email');
        
        $config['protocol'] = $this->protocolo;
        $config['smtp_host'] = $this->servidor_correo;
        $config['smtp_timeout'] = '10';                     
        $config['smtp_user'] = $this->usuario_sistema;     
        $config['smtp_pass'] = $this->password_sistema;              
        $config['mailtype'] = 'html'; 
        $this->email->initialize($config);
        
        //Preparando el mensaje
        $this->email->from($this->correo_sistema, 'Gestión de Contratos - Hatovial S.A.S.');
        $this->email->to($this->usuarios_correo); 
        $this->email->cc(''); 
        //$this->email->bcc('johnarleycano@hotmail.com'); 
        $this->email->subject($asunto);                     
        
        $mensaje = file_get_contents('application/views/email/plantilla.html');
        $mensaje = str_replace('{TITULO}', 'Hatovial S.A.S. - Gestión de Contratos', $mensaje);
        $mensaje = str_replace('{MENSAJE}', $cuerpo, $mensaje) ;
        
        $this->email->message($mensaje);	
        $this->email->send();
    }//Fin enviar()

    /**
    * Env&iacute;a los correos electr&oacute;nicos
    * 
    * @access	private
    */
    function contratos_en_vencimiento(){
        $sql = 
        'SELECT
            contratos.Numero,
            tbl_terceros.Nombre AS Contratista,
            contratos.Objeto,
            contratos.Fecha_Vencimiento,
            contratos.Fecha_Vencimiento - CURDATE() AS Dias_Restantes
        FROM
            contratos
        INNER JOIN tbl_terceros ON contratos.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
        WHERE (contratos.Fecha_Vencimiento - CURDATE()) BETWEEN 0 AND 5
        ORDER BY
            contratos.Fecha_Vencimiento ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin contratos_en_vencimiento

    /**
    * Env&iacute;a los correos electr&oacute;nicos
    * 
    * @access	private
    */
    function contratos_vencidos(){
        $sql = 
        'SELECT
            contratos.Numero,
            contratos.Objeto,
            tbl_terceros.Nombre AS Contratista,
            contratos.Fecha_Vencimiento,
            DATEDIFF(CURDATE(),contratos.Fecha_Vencimiento) AS Dias_Vencidos
        FROM
            contratos
        INNER JOIN tbl_terceros ON contratos.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
        WHERE
            contratos.Fecha_Vencimiento < CURDATE() AND
            contratos.Fk_Id_Estado != 2
        ORDER BY
            contratos.Numero ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin contratos_vencidos
    
    /**
    * Env&iacute;a los correos electr&oacute;nicos
    * 
    * @access	private
    */
    function polizas_en_vencimiento(){
        $sql = 
        'SELECT
        contratos.Numero,
        contratos.Objeto,
        tbl_terceros.Nombre AS Contratista,
        poliza_tipo.Poliza_Tipo,
        contratos_polizas.Fecha_Final,
        DATEDIFF(contratos_polizas.Fecha_Final, CURDATE()) AS Dias_Vencidos
        FROM
        contratos_polizas
        INNER JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        INNER JOIN contratos ON contratos_polizas.Fk_Id_Contratos = contratos.Pk_Id_Contrato
        INNER JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos.Fk_Id_Terceros
        WHERE
        (contratos_polizas.Fecha_Final - CURDATE()) BETWEEN 0 AND 5
        ORDER BY
        contratos.Numero ASC,
        poliza_tipo.Pk_Id_Poliza_Tipo ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin contratos_en_vencimiento
    
    /**
    * Env&iacute;a los correos electr&oacute;nicos
    * 
    * @access	private
    */
    function polizas_vencidas(){
        $sql = 
        'SELECT
        contratos.Numero,
        contratos.Objeto,
        tbl_terceros.Nombre AS Contratista,
        poliza_tipo.Poliza_Tipo,
        contratos_polizas.Fecha_Final,
        DATEDIFF(CURDATE(), contratos_polizas.Fecha_Final) AS Dias_Vencidos
        FROM
        contratos_polizas
        INNER JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        INNER JOIN contratos ON contratos_polizas.Fk_Id_Contratos = contratos.Pk_Id_Contrato
        INNER JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos.Fk_Id_Terceros
        WHERE
        contratos_polizas.Fecha_Final < CURDATE() AND
        contratos.Fk_Id_Estado != 2 AND
        contratos_polizas.Fecha_Final != "0000-00-00"
        ORDER BY
        contratos.Numero ASC,
        poliza_tipo.Pk_Id_Poliza_Tipo ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin polizas_vencidas()
    
    /**
    * Env&iacute;a los correos electr&oacute;nicos
    * 
    * @access	private
    */
    function no_acta_inicio(){
        $sql = 
        'SELECT
            contratos.Numero,
            contratos.Objeto,
            tbl_terceros.Nombre AS Contratista,
            contratos.Fecha_Inicial,
            contratos.Fecha_Vencimiento,
            contratos.Valor_Inicial,
            tbl_estados.Estado
        FROM
            contratos
            INNER JOIN tbl_terceros ON contratos.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
            INNER JOIN tbl_estados ON contratos.Fk_Id_Estado = tbl_estados.Pk_Id_Estado
        WHERE
            contratos.Acta_Inicio IS FALSE AND
            tbl_estados.Pk_Id_Estado <> 2
        ORDER BY
            contratos.Numero ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin no_acta_inicio()
}//Fin email
/* End of file email_model.php */
/* Location: ./contratos/application/controllers/email_model.php */