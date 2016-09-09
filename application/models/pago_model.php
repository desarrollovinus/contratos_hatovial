<?php
/**
 * Modelo que se encarga de gestionar los pagos
 * a los contratos.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Pago_model extends CI_Model{
    
    /**
    * Guarda en la base de datos los pagos al contrato.
    *
    * @access	public
    * @return	
    */
    function agregar_pago($pago){
        //Se insertan los datos del pago
        $this->db->insert('contratos_pagos', $pago);
    }//Fin agregar_pago()
    
    /**
    * Elimina un pago de la base de datos.
    *
    * @access	public
    * @return	
    */
    function eliminar_pago($id_pago){
        $this->db->delete('contratos_pagos', array('Pk_Id_Contratos_Pagos' => $id_pago));
    }
    
    /**
    * Trae los pagos asociados a un contrato.
    *
    * @access	public
    * @return	
    */
    function listar_pagos($id_contrato){
        //Consulta
        $sql = 
        'SELECT
            contratos_pagos.Pk_Id_Contratos_Pagos,
            contratos.Pk_Id_Contrato,
            contratos.Numero,
            contratos_pagos.Fecha,
            contratos_pagos.Numero_Acta,
            contratos_pagos.Valor_Pago,
            contratos_pagos.Valor_Retenido
        FROM
            contratos
        INNER JOIN contratos_pagos ON contratos_pagos.Fk_Id_Contratos = contratos.Pk_Id_Contrato
        WHERE
            contratos.Pk_Id_Contrato = '.$id_contrato;
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_pagos
    
    /**
    * Trae el estado de los pagos de un contrato.
    *
    * @access	public
    * @return	
    */
    function estado_pagos_contrato($id_contrato){
        //Consulta
        $sql =
        'SELECT
        contratos.Numero,
        contratos.Valor_Inicial,
        (SELECT
        IFNULL(SUM(contratos_adiciones.Valor), 0)
        FROM
        contratos_adiciones
        WHERE
        contratos_adiciones.Fk_Id_Contrato = contratos.Pk_Id_Contrato) AS Valor_Adiciones,
        IFNULL(sum(contratos_pagos.Valor_Retenido), 0) AS Valor_Retenido,
        IFNULL(sum(contratos_pagos.Valor_Pago), 0) AS Pagado,
        IFNULL(((IFNULL(sum(contratos_pagos.Valor_Pago), 0))*100)/(contratos.Valor_Inicial + (SELECT
        IFNULL(SUM(contratos_adiciones.Valor), 0)
        FROM
        contratos_adiciones
        WHERE
        contratos_adiciones.Fk_Id_Contrato = contratos.Pk_Id_Contrato)),100) AS Porcentaje
        FROM
        contratos
        INNER JOIN contratos_pagos ON contratos_pagos.Fk_Id_Contratos = contratos.Pk_Id_Contrato
        WHERE
        contratos.Pk_Id_Contrato =  '.$id_contrato;
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin estado_pagos_contrato()
    
    function pagos_excedidos(){
        //Consulta
        $sql =
        /*'SELECT
        contratos.Numero,
        contratos.Objeto,
        tbl_terceros.Nombre AS Contratista,
        contratos.Valor_Inicial,
        Sum(contratos_pagos.Valor_Pago) AS Pagado,
        (SUM(contratos_pagos.Valor_Pago) - contratos.Valor_Inicial) AS Excedido
        FROM
        contratos
        INNER JOIN contratos_pagos ON contratos_pagos.Fk_Id_Contratos = contratos.Pk_Id_Contrato
        INNER JOIN tbl_terceros ON contratos.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
        GROUP BY
        contratos.Numero
        ORDER BY
        contratos.Numero ASC';*/
        'SELECT
            contratos.Pk_Id_Contrato,
            contratos.Numero,
            contratos.Objeto,
            tbl_terceros.Nombre AS Contratista,
            contratos.Valor_Inicial,
            (SELECT
                IFNULL(SUM(contratos_adiciones.Valor), 0)
            FROM
                contratos_adiciones
            WHERE
                contratos_adiciones.Fk_Id_Contrato = contratos.Pk_Id_Contrato) AS Adiciones,
            contratos.Valor_Inicial + 
            (SELECT
                IFNULL(SUM(contratos_adiciones.Valor), 0)
            FROM
                contratos_adiciones
            WHERE
                contratos_adiciones.Fk_Id_Contrato = contratos.Pk_Id_Contrato) AS Valor_Total,
            Sum(contratos_pagos.Valor_Pago) AS Pagado,
            (SUM(contratos_pagos.Valor_Pago) - (contratos.Valor_Inicial + 
            (SELECT
                IFNULL(SUM(contratos_adiciones.Valor), 0)
            FROM
                contratos_adiciones
            WHERE
                contratos_adiciones.Fk_Id_Contrato = contratos.Pk_Id_Contrato))) AS Excedido
            FROM
                contratos
                INNER JOIN contratos_pagos ON contratos_pagos.Fk_Id_Contratos = contratos.Pk_Id_Contrato
                INNER JOIN tbl_terceros ON contratos.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
            WHERE
                contratos.Fk_Id_Estado <> 2
            GROUP BY
                contratos.Numero
            ORDER BY
                contratos.Numero ASC';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Finpagos_excedidos()
}//Fin pago_model