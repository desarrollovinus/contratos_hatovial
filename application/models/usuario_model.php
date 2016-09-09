<?php
/**
 * Modelo que se encarga de gestionar la informaci&oacute;n
 * de los usuarios del sistema.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Usuario_model extends CI_Model{
    
    /**
    * Retorna los datos de un usuario si existe.
    *
    * @access	public
    * @param	string	el usuario.
    * @param	string	el password.
    * @return	los datos del usuario, FALSE en caso contrario.
    */
    function validar_login($usuario, $password){
        //columnas que se van a retornar
        $this->db->select('Pk_Id_Usuario');
        $this->db->select('Nombres');
        $this->db->select('Apellidos');
        $this->db->select('Usuario');
        $this->db->select('Email');
        $this->db->select('Tipo');
        $this->db->select('Activo');
        
        //Se valida el usuario y el password ingresados
        $this->db->where('Usuario', $usuario);
        $this->db->where('Clave', $password);
        
        //Se retorna la consulta
        $resultado = $this->db->get('tbl_usuarios')->row();
        
        return $resultado;
    }//Fin validar_login()
    
    /**
    * Retorna verdadero si el nombre de usuario que se ingresa existe.
    *
    * @access	public
    * @param	string	el usuario.
    * @return	los datos del nombre de usuario, TRUE en caso contrario.
    */
    function validar_usuario($usuario){
        $this->db->where('Usuario', $usuario);
        $query = $this->db->get('tbl_usuarios');
        if($query->num_rows() >0 ){
            return false;
        }else{
            return true;
        }
    }//Fin validar_usuario()
    
    /**
    * Retorna verdadero si el correo electr&oacute;nico que se ingresa existe.
    *
    * @access	public
    * @param	string	el usuario.
    * @return	el correo electr&oacute;nico del usuario, TRUE en caso contrario.
    */
    function validar_email($email){
        $this->db->where('Email', $email);
        $query = $this->db->get('tbl_usuarios');
        if($query->num_rows() >0 ){
            return false;
        }else{
            return true;
        }
    }//Fin validar_email()
    
    /**
    * Inserta un usuario en la base de datos
    *
    * @access	public
    */
    function insertar_usuario($usuario_nuevo){
        $this->db->insert('tbl_usuarios', $usuario_nuevo);
    }//Fin insertar_usuario()
}
    