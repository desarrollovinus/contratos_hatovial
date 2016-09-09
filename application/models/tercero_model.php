<?php
/**
 * Modelo que se encarga de gestionar la informaci&oacute;n
 * de los terceros.
 * @author 		John Arley Cano Salinas
 * @copyright           &copy; HATOVIAL S.A.S.
 */
Class Tercero_model extends CI_Model{
    /**
    * Guarda en la base de datos los datos del contrato.
    *
    * @access	public
    * @return	
    */
    function registrar_tercero($tercero){
        //Se insertan los datos del tercero
        $this->db->insert('tbl_terceros', $tercero);
    }//Fin registrar_tercero
    
    /**
    * Actualiza en la base de datos los datos del contrato.
    *
    * @access	public
    * @return	
    */
    function actualizar_tercero($tercero, $id_tercero){
        //Se actualizan los datos del tercero
        $this->db->where('Pk_Id_Terceros', $id_tercero);
        $this->db->update('tbl_terceros', $tercero);
    }//Fin actualizar_tercero()
    
    /**
    * Lista los tipos de terceros existentes en la base de datos.
    *
    * @access	public
    * @return	
    */
    function listar_tipo_terceros(){
        //columnas que se van a retornar
        $this->db->select('*');
        //$this->db->where('Fk_Id_Terceros_Tipo', 2);
        $this->db->order_by('Nombre', "asc"); 
        
        //Se retorna la consulta
        $resultado = $this->db->get('tbl_terceros_tipo')->result();
        
        return $resultado;
    }//Fin listar_tipo_terceros()
    
    /**
    * Lista los terceros existentes en la base de datos.
    *
    * @access	public
    * @return	
    */
    function listar_terceros($id_tercero){
        //columnas que se van a retornar
        if(isset($id_tercero)){$filtro = ' Where Pk_Id_Terceros = '.$id_tercero;}else{$filtro = '';}
        $sql = 
        'SELECT
        tbl_terceros_tipo.Nombre AS Tipo,
        tbl_terceros.Pk_Id_Terceros,
        tbl_terceros.Fk_Id_Terceros_Tipo,
        tbl_terceros.Documento,
        tbl_terceros.Nombre,
        tbl_terceros.Telefono,
        tbl_terceros.Direccion,
        tbl_terceros.Representante_Legal
        FROM
        tbl_terceros
        INNER JOIN tbl_terceros_tipo ON tbl_terceros_tipo.Pk_Id_Terceros_Tipo = tbl_terceros.Fk_Id_Terceros_Tipo '.$filtro;
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_terceros()
    
    /**
    * Retorna los contratantes registrados en el sistema.
    *
    * @access   
    * @param    
    * @param    
    * @return   
    */
    function listar_contratantes(){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->where('Fk_Id_Terceros_Tipo', 3);
        $this->db->order_by('Nombre', "asc"); 
        
        //Se retorna la consulta
        $resultado = $this->db->get('tbl_terceros')->result();
        
        return $resultado;
    }//Fin listar_contratantes()
    
    /**
    * Retorna los contratistas registrados en el sistema.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_contratistas(){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->where('Fk_Id_Terceros_Tipo', 2);
        $this->db->order_by('Nombre', "asc"); 
        
        //Se retorna la consulta
        $resultado = $this->db->get('tbl_terceros')->result();
        
        return $resultado;
    }//Fin listar_contratistas()
    
    /**
    * Retorna las aseguradoras para las polizas de un contrato.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_aseguradoras(){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->where('Fk_Id_Terceros_Tipo', 1);
        $this->db->order_by('Nombre', "asc");
        
        //Se retorna la consulta
        $resultado = $this->db->get('tbl_terceros')->result();
        
        return $resultado;
    }//Fin listar_aseguradoras()
}//Fin tercero_model