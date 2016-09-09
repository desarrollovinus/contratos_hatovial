<?php
/**
 * Modelo que se encarga de gestionar la informaci&oacute;n
 * de los contratos en estado de liquidaci&oacute;n.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Liquidacion_model extends CI_Model{
    /**
    * Guarda en la base de datos los datos del contrato.
    *
    * @access	public
    * @return	
    */
    function listar_contratos(){
        //Consulta
        $sql = 
        'SELECT
            contratos.Pk_Id_Contrato,
            contratos.Numero,
            tbl_terceros.Nombre AS Contratista,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Documento AS Documento_Contratista,
            tbl_terceros.Telefono AS Telefono_Contratista,
            tbl_terceros.Direccion AS Direccion_Contratista,
            contratos.Objeto,
            contratos.Lugar,
            tbl_estados.Pk_Id_Estado,
            tbl_estados.Estado,
            contratos.Valor_Inicial,
            contratos.Fecha_Inicial,
            contratos.Plazo,
            contratos.Acta_Inicio,
            contratos.Fecha_Acta_Inicio
        FROM
            contratos
        LEFT JOIN tbl_estados ON (contratos.Fk_Id_Estado = tbl_estados.Pk_Id_Estado)
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos.Fk_Id_Terceros
        WHERE
        tbl_estados.Pk_Id_Estado = 2';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin listar_contratos
}//Fin liquidacion_model