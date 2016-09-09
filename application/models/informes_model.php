<?php
/**
 * Modelo que se encarga de gestionar la informaci&oacute;n
 * de las aseguradoras.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Informes_model extends CI_Model{
    /**
    * Lista los contratos seg&uacute;n el estado seleccionado
    *
    * @access	public
    */
    function contratos_por_estado($id_estado){
        $sql =
        'SELECT
        contratos.contratos.Numero,
        contratos.contratos.Objeto,
        contratos.contratos.Lugar,
        contratos.contratos.Valor_Inicial,
        contratos.contratos.Acta_Inicio,
        contratos.contratos.Fecha_Acta_Inicio,
        contratos.contratos.Plazo,
        contratos.contratos.Fecha_Vencimiento,
        contratos.contratos.Fecha_Inicial,
        contratos.tbl_terceros.Nombre AS Contratista,
        contratos.tbl_estados.Estado
        FROM
        contratos.contratos
        INNER JOIN contratos.tbl_terceros ON contratos.contratos.Fk_Id_Terceros = contratos.tbl_terceros.Pk_Id_Terceros
        INNER JOIN contratos.tbl_estados ON contratos.contratos.Fk_Id_Estado = contratos.tbl_estados.Pk_Id_Estado
        WHERE
        contratos.contratos.Fk_Id_Estado = '.$id_estado.' ORDER BY contratos.contratos.Numero ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin contratos_por_estado()
    
    /**
    * Lista los contratos seg&uacute;n el rango de fechas seleccionado
    *
    * @access	public
    */
    function contratos_por_fecha_inicial($fecha, $fecha1, $fecha2){
        $sql =
        "SELECT
            contratos.contratos.Numero,
            contratos.contratos.Objeto,
            contratos.contratos.Lugar,
            contratos.contratos.Acta_Inicio,
            contratos.contratos.Fecha_Acta_Inicio,
            contratos.contratos.Fecha_Vencimiento,
            contratos.contratos.Plazo,
            contratos.contratos.Valor_Inicial,
            contratos.tbl_estados.Estado,
            contratos.tbl_terceros.Nombre AS Contratista,
            contratos.contratos.Fecha_Inicial
        FROM
            contratos.contratos
            INNER JOIN contratos.tbl_terceros ON contratos.contratos.Fk_Id_Terceros = contratos.tbl_terceros.Pk_Id_Terceros
            INNER JOIN contratos.tbl_estados ON contratos.contratos.Fk_Id_Estado = contratos.tbl_estados.Pk_Id_Estado
        WHERE
            contratos.".$fecha." BETWEEN '".$fecha1."' AND '".$fecha2."
        ORDER BY
            contratos.contratos.Numero ASC'";
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin contratos_por_fecha_inicial()

    /**
    * Lista los contratos con sus respectivos valores
    *
    * @access   public
    */
    function listar_valores($tipo, $id){
        $contratista = "";
        $contratante = "";
        
        // Si viene algun dato y es contratista
        if ($tipo == "contratista" && $id != "") {
            $contratista = "WHERE c.Fk_Id_Terceros = {$id}";
        }
        
        // Si viene algun dato y es contratista
        if ($tipo == "contratante" && $id != "") {
            $contratista = "WHERE c.Fk_Id_Terceros_Contratante = {$id}";
        }

        $sql =
        "SELECT
            c.Pk_Id_Contrato,
            c.Numero,
            c.Objeto,
            tc.Nombre AS Contratante,
            tbl_terceros.Nombre AS Contratista,
            c.Fecha_Inicial,
            c.Fecha_Vencimiento,
            c.Plazo AS Plazo_Inicial,
            c.Valor_Inicial,
            (
                SELECT
                    ifnull(SUM(contratos_pagos.Valor_Pago), 0)
                FROM
                    contratos
                LEFT JOIN contratos_pagos ON contratos_pagos.Fk_Id_Contratos = contratos.Pk_Id_Contrato
                WHERE
                    contratos.Pk_Id_Contrato = c.Pk_Id_Contrato
                GROUP BY
                    contratos.Pk_Id_Contrato
            ) AS Pagado,
            (
                SELECT
                    IFNULL(SUM(adiciones.Plazo), 0)
                FROM
                    contratos_adiciones AS adiciones
                WHERE
                    adiciones.Fk_Id_Contrato = c.Pk_Id_Contrato
            ) AS Plazo_Adiciones,
            (
                SELECT
                    IFNULL(
                        SUM(contratos_adiciones.Valor),
                        0
                    )
                FROM
                    contratos_adiciones
                WHERE
                    contratos_adiciones.Fk_Id_Contrato = c.Pk_Id_Contrato
            ) AS Valor_Adiciones,
            tbl_estados.Estado
        FROM
            contratos AS c
        LEFT JOIN tbl_terceros ON c.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
        LEFT JOIN tbl_estados ON c.Fk_Id_Estado = tbl_estados.Pk_Id_Estado
        LEFT JOIN tbl_terceros AS tc ON c.Fk_Id_Terceros_Contratante = tc.Pk_Id_Terceros
        {$contratista}
        ORDER BY
            c.Fk_Id_Estado ASC,
            c.Fecha_Inicial DESC";

        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_valores()

    /**
    * Lista los usuarios que tienen contratos creados
    *
    * @access   public
    */
    function listar_usuarios(){
        //Se filtran solo id los diferentes 
        $this->db->distinct();

        //Se seleccionan las columnas
        $this->db->select('Fk_Id_Usuario');
        $this->db->select('Nombres');
        $this->db->select('Apellidos');
        $this->db->order_by('Nombres', 'asc');
        
        $this->db->from('contratos');
        $this->db->join('tbl_usuarios', 'contratos.Fk_Id_Usuario = tbl_usuarios.Pk_Id_Usuario');
        return $this->db->get()->result();
    }//Fin listar_usuarios()

    function obtener_nombre_contratista($id_contratista){
        $this->db->select('Nombre');
        $this->db->where('Pk_Id_Terceros', $id_contratista);

        return $this->db->get('tbl_terceros')->row();
    }//Fin obtener_nombre_contratante

    /**
    * Retorna el nombre del usuario especifico
    *
    * @access   public
    */
    function obtener_nombre_usuario($id_usuario){
        $this->db->select('*');
        $this->db->where('Pk_Id_Usuario', $id_usuario);

        foreach ($this->db->get('tbl_usuarios')->result() as $usuario) {
            return $usuario->Nombres.' '.$usuario->Apellidos;
        }
    }//Fin obtener_nombre_usuario

    /**
    * Cuenta los estados que hay por cada contrato
    *
    * @access   public
    */
    function contar_estados(){
        //Se seleccionan la consulta
        $sql =
        'SELECT
        tbl_estados.Estado,
        COUNT(contratos.Pk_Id_Contrato) AS Total
        FROM
        tbl_estados
        INNER JOIN contratos ON contratos.Fk_Id_Estado = tbl_estados.Pk_Id_Estado
        GROUP BY
        contratos.Fk_Id_Estado';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin contar_estados()
}//Fin informes_model
