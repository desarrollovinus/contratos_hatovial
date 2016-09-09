<?php
/**
 * Modelo que se encarga de guardar un registro en la base de datos
 * de las anotaciones de la bit&aacute;cora para cada contrato
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Bitacora_model extends CI_Model{
    
    /**
    * Guarda en la base de datos los datos de la bit&aacute;cora.
    *
    * @access	public
    * @return	
    */
    function agregar_bitacora($bitacora){
        //Se insertan los datos principales del contrato
        $this->db->insert('contratos_bitacora', $bitacora);
    }//Fin agregar_bitacora
    
    /**
    * Elimina de la base de datos los datos la anotaci&oacute;n
    * de la bit&aacute;cora seleccionada.
    *
    * @access	public
    * @return	
    */
    function eliminar_bitacora($bitacora){
        //Se insertan los datos principales del contrato
        $this->db->insert('contratos_bitacora', $bitacora);
    }//Fin agregar_bitacora
    
    /**
    * Retorna los datos de la bit&aacute;cora del contrato
     * para la tabla de la vista.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_bitacora($id_contrato){
        //Consulta
        $sql = 
        'SELECT
        DATE_FORMAT(contratos_bitacora.Fecha,"%d-%m-%Y") AS Fecha,
            contratos_bitacora.Asunto,
            contratos_bitacora.Observacion,
            contratos_bitacora.Fk_Id_Contratos,
            contratos.Numero,
            tbl_usuarios.Nombres,
            tbl_usuarios.Apellidos
        FROM
            contratos_bitacora
            LEFT JOIN contratos ON contratos_bitacora.Fk_Id_Contratos = contratos.Pk_Id_Contrato
            INNER JOIN tbl_usuarios ON contratos_bitacora.Fk_Id_Usuario = tbl_usuarios.Pk_Id_Usuario
        WHERE
            contratos_bitacora.Fk_Id_Contratos = '.$id_contrato;
        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin listar_contratos
}//Fin bitacora_model
