<?php
/**
 * Modelo que se encarga de guardar un registro en la base de datos
 * de las acciones que se realizan en la aplicaci&oacute;n
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Auditoria_model extends CI_Model{
    
    /**
    * Agrega la auditor&iacute;a cuando un usuario intenta ingresar
     * a la aplicaci&oacute;n.
    *
    * @access	public
    */
    function intento_inicio_sesion($usuario){
        $auditoria = array(
            'descripcion' => $usuario.' intenta validarse',
            'Fk_Id_Auditoria_Tipo' => 1
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin intento_inicio_sesion()
    
    /**
    * Agrega la auditor&iacute;a cuando un usuario inicia correctamente sesi&oacute;n.
    *
    * @access	public
    */
    function iniciar_sesion(){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Ingresa a la aplicación',
            'Fk_Id_Auditoria_Tipo' => 2
        );
        $this->db->insert('auditoria', $auditoria);
        //Fin de auditoria
    }//Fin iniciar_sesion()
    
    /**
    * Agrega la auditor&iacute;a cuando se cierra la sesi&oacute;n.
    *
    * @access	public
    */
    function cerrar_sesion(){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Sale de la aplicación',
            'Fk_Id_Auditoria_Tipo' => 3
        );
        $this->db->insert('auditoria', $auditoria);
        //Fin de auditoria
    }//Fin  cerrar_sesion
    
    /**
    * Agrega la auditor&iacute;a cuando se inserta un nuevo usuario.
    *
    * @access	public
    */
    function insertar_usuario($usuario, $id_usuario){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $id_usuario,
            'descripcion' => 'Crea el usuario '.$usuario,
            'Fk_Id_Auditoria_Tipo' => 4
        );
        $this->db->insert('auditoria', $auditoria);
        //Fin de auditoria
    }//Fin insertar_usuario
    
    /**
    * Agrega la auditor&iacute;a cuando se inserta un contrato nuevo.
    *
    * @access	public
    */
    function insertar_contrato($numero){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Crea el contrato '.$numero,
            'Fk_Id_Auditoria_Tipo' => 5
        );
        $this->db->insert('auditoria', $auditoria);
        //Fin de auditoria
    }//Fin insertar_usuario
    
    /**
    * Agrega la auditor&iacute;a cuando se modifica un contrato existente.
    *
    * @access	public
    */
    function modificar_contrato($numero){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Modifica el contrato '.$numero,
            'Fk_Id_Auditoria_Tipo' => 6
        );
        $this->db->insert('auditoria', $auditoria);
        //Fin de auditoria
    }//Fin modificar_contrato
    
    /**
    * Agrega la auditor&iacute;a cuando se crea un tercero nuevo.
    *
    * @access	public
    */
    function insertar_tercero($id_tercero, $nombre){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Crea el tercero '.$id_tercero.' ('.$nombre.')',
            'Fk_Id_Auditoria_Tipo' => 7
        );
        
        $this->db->insert('auditoria', $auditoria);
        
    }//Fin insertar_tercero
    
    /**
    * Agrega la auditor&iacute;a cuando se modifica un tercero.
    *
    * @access	public
    */
    function modificar_tercero($id_tercero, $nombre){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Modifica el tercero '.$id_tercero.' ('.$nombre.')',
            'Fk_Id_Auditoria_Tipo' => 8
        );
        
        $this->db->insert('auditoria', $auditoria);
        
    }//Fin modificar_tercero
    
    /**
    * Agrega la auditor&iacute;a cuando se crea un pago a un contrato.
    *
    * @access	public
    */
    function agregar_pago($valor_pago, $numero){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Realiza un pago por $'.number_format($valor_pago, 0, '', '.') .' al contrato '.$numero,
            'Fk_Id_Auditoria_Tipo' => 9
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin agregar_pago()
    
    /**
    * Agrega la auditor&iacute;a cuando se crea un pago a un contrato.
    *
    * @access	public
    */
    function eliminar_pago($valor, $numero){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Elimina un pago por $'.number_format($valor, 0, '', '.').' al contrato '.$numero,
            'Fk_Id_Auditoria_Tipo' => 10
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin agregar_pago()
    
    /**
    * Agrega la auditor&iacute;a cuando se crea una anortaci&oacute;n
    * a la bit&aacute;cora de un contrato.
    *
    * @access	public
    */
    function agregar_bitacora($numero){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'descripcion' => 'Agrega un registro a la bitácora del contrato '.$numero,
            'Fk_Id_Auditoria_Tipo' => 11
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin agregar_bitacora()
    
    /**
    * Agrega la auditor&iacute;a cuando se sube un archivo a la aplicaci&oacute;n
    *
    * @access	public
    */
    function subir_archivo($archivo){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'Descripcion' => 'Sube el archivo '.$archivo,
            'Fk_Id_Auditoria_Tipo' => 14
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin agregar_bitacora()
    
    /**
    * Lista todas las acciones de auditor&iacute;a del sistema.
    *
    * @access	public
    */
    function listar_auditoria(){
        $sql =
        "SELECT
            DATE_FORMAT(auditoria.Fecha_Hora,'%d-%m-%Y') AS Fecha,
            DATE_FORMAT(auditoria.Fecha_Hora,'%h:%i:%s %p') AS Hora,
            IFNULL(concat(tbl_usuarios.Nombres, ' ', tbl_usuarios.Apellidos),'') AS Nombres,
            IFNULL(tbl_usuarios.Usuario, '') AS Usuario,
            auditoria.Descripcion       
        FROM
            auditoria
            LEFT OUTER JOIN tbl_usuarios ON tbl_usuarios.Pk_Id_Usuario = auditoria.Fk_Id_Usuario
        ORDER BY
            auditoria.Id_Auditoria DESC";
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_auditoria()
    
    /**
    * Formatea las fechas de manera que salgan los meses y d&iacute;s alfab&eacute;ticos
    *
    * @access	public
    */
    function formato_fecha($fecha){
        //Si No hay fecha, devuelva vac&iacute;o en vez de 0000-00-00
        if($fecha == '0000-00-00' || $fecha == ''){
            return false;
        }
        
        $dia_num = date("j", strtotime($fecha));
        $dia = date("N", strtotime($fecha));
        $mes = date("m", strtotime($fecha));
        $anio_es = date("Y", strtotime($fecha));

        //Nombres de los d&iacute;as
        if($dia == "1"){ $dia_es = "Lunes"; }
        if($dia == "2"){ $dia_es = "Martes"; }
        if($dia == "3"){ $dia_es = "Miercoles"; }
        if($dia == "4"){ $dia_es = "Jueves"; }
        if($dia == "5"){ $dia_es = "Viernes"; }
        if($dia == "6"){ $dia_es = "Sabado"; }
        if($dia == "7"){ $dia_es = "Domingo"; }

        //Nombres de los meses
        if($mes == "1"){ $mes_es = "enero"; }
        if($mes == "2"){ $mes_es = "febrero"; }
        if($mes == "3"){ $mes_es = "marzo"; }
        if($mes == "4"){ $mes_es = "abril"; }
        if($mes == "5"){ $mes_es = "mayo"; }
        if($mes == "6"){ $mes_es = "junio"; }
        if($mes == "7"){ $mes_es = "julio"; }
        if($mes == "8"){ $mes_es = "agosto"; }
        if($mes == "9"){ $mes_es = "septiembre"; }
        if($mes == "10"){ $mes_es = "octubre"; }
        if($mes == "11"){ $mes_es = "noviembre"; }
        if($mes == "12"){ $mes_es = "diciembre"; } 

        //a&ntilde;o
        //$anio_es = $anio_es;

        //Se foramtea la fecha
        $fecha = /*$dia_es." ".*/$dia_num." de ".$mes_es." de ".$anio_es;
        
        return $fecha;
    }//Fin formato_fecha()

    /**
    * Genera el acta de inicio
    *
    * @access   public
    */
    function generar_acta_inicio($id_contrato){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'Descripcion' => 'Genera el acta de inicio para el contrato '.$id_contrato,
            'Fk_Id_Auditoria_Tipo' => 15
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin generar_acta_inicio()

    /**
    * Genera el acta de recibo
    *
    * @access   public
    */
    function generar_acta_recibo($id_contrato){
        //Accion de auditoria
        $auditoria = array(
            'Fk_Id_usuario' => $this->session->userdata('Pk_Id_Usuario'),
            'Descripcion' => 'Genera el acta de recibo para el contrato '.$id_contrato,
            'Fk_Id_Auditoria_Tipo' => 16
        );
        $this->db->insert('auditoria', $auditoria);
    }//Fin generar_acta_recibo()
}//Fin auditoria_model
/* End of file auditoria_model.php */
/* Location: ./contratos/application/controllers/auditoria_model.php */