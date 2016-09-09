<?php
/**
 * Modelo que se encarga de gestionar la informaci&oacute;n
 * de los contratos.
 * @author 		John Arley Cano Salinas
 * @copyright	&copy; HATOVIAL S.A.S.
 */
Class Contrato_model extends CI_Model{
    
    /**
    * Guarda en la base de datos los datos del contrato.
    *
    * @access	public
    * @return	
    */
    function registrar_contrato($contrato){
        //Se insertan los datos principales del contrato
        $this->db->insert('contratos', $contrato);
    }//Fin registrar_contrato
    
    /**
    * Actualiza en la base de datos los datos del contrato.
    *
    * @access	public
    * @return	
    */
    function actualizar_contrato($contrato, $id_contrato){
        //Se actualizan los datos principales del contrato
        $this->db->where('Pk_Id_Contrato', $id_contrato);
        $this->db->update('contratos', $contrato);
    }//Fin registrar_contrato
    
    /**
    * Retorna verdadero si el numero de contrato que se ingresa existe.
    *
    * @access	public
    * @param	string	el numero del contrato.
    * @return	
    */
    function verificar_numero($numero){
        $this->db->where('Numero', $numero);
        $query = $this->db->get('contratos');
        if($query->num_rows() >0 ){
            return false;
        }else{
            return true;
        }
    }//Fin verificar_numero()
    
    /**
    * Guarda en la base de datos de las polizas relacionadas al contrato.
    *
    * @access	public
    * @return	
    */
    function registrar_polizas($poliza_cumplimiento, $poliza_prestaciones, $poliza_anticipos, $poliza_calidad, $poliza_estabilidad, $poliza_rc){
        //Se inserta la poliza de cumplimiento
        $this->db->insert('contratos_polizas', $poliza_cumplimiento);
        $this->db->insert('contratos_polizas', $poliza_prestaciones);
        $this->db->insert('contratos_polizas', $poliza_anticipos);
        $this->db->insert('contratos_polizas', $poliza_calidad);
        $this->db->insert('contratos_polizas', $poliza_estabilidad);
        $this->db->insert('contratos_polizas', $poliza_rc);
    }//Fin registrar_polizas()
    
    /**
    * Actualiza en la base de datos los datos de la poliza de cumplimiento.
    *
    * @access	public
    * @return	
    */
    function actualizar_polizas(
        $poliza_cumplimiento, $id_cumplimiento, 
        $poliza_prestaciones, $id_prestaciones, 
        $poliza_anticipos, $id_anticipos, $poliza_calidad, 
        $id_calidad, $poliza_estabilidad, $id_estabilidad,
        $poliza_rc, $id_rc){
            //Se actualizan los datos de las polizas
            $this->db->where('Pk_Id_Contratos_Poliza', $id_cumplimiento);
            $this->db->update('contratos_polizas', $poliza_cumplimiento);

            $this->db->where('Pk_Id_Contratos_Poliza', $id_prestaciones);
            $this->db->update('contratos_polizas', $poliza_prestaciones);

            $this->db->where('Pk_Id_Contratos_Poliza', $id_anticipos);
            $this->db->update('contratos_polizas', $poliza_anticipos);

            $this->db->where('Pk_Id_Contratos_Poliza', $id_calidad);
            $this->db->update('contratos_polizas', $poliza_calidad);

            $this->db->where('Pk_Id_Contratos_Poliza', $id_estabilidad);
            $this->db->update('contratos_polizas', $poliza_estabilidad);

            $this->db->where('Pk_Id_Contratos_Poliza', $id_rc);
            $this->db->update('contratos_polizas', $poliza_rc);
    }//Fin actualizar_polizas
    
    /**
    * Retorna los estados de un contrato.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_contratos_estados(){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->select('Estado');
        
        //Se retorna la consulta
        $resultado = $this->db->get('tbl_estados')->result();
        
        return $resultado;
    }//Function listar_contratos_estados()
    
    /**
    * Retorna el id del contrato que se acaba de crear.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function obtener_id_contrato(){
        return @mysql_insert_id($this->Pk_Id_Contrato);
    }//Fin obtener_id_contrato()
    
    /**
    * Retorna los datos del contrato para la tabla de la vista de inicio.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_contratos($id_contrato){
        //Esta validaci&oacute;n verifica si llega algun valor en el id, con el
        //fin de filtrar la consulta por un solo id de contrato
        if($id_contrato){
            $id_contrato = 'WHERE c.Pk_Id_Contrato = '.$id_contrato;
        }
        
        //Consulta
        $sql = 
        'SELECT
            c.Pk_Id_Contrato,
            c.Numero,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre AS Contratista,
            tbl_terceros.Documento AS Documento_Contratista,
            tbl_terceros.Telefono AS Telefono_Contratista,
            tbl_terceros.Direccion AS Direccion_Contratista,
            tbl_terceros.Representante_Legal,
            c.Objeto,
            c.Lugar,
            tbl_estados.Pk_Id_Estado,
            tbl_estados.Estado,
            c.Valor_Inicial,
            (SELECT
                IFNULL(SUM(contratos_adiciones.Valor), 0)
            FROM
                contratos_adiciones
            WHERE
                contratos_adiciones.Fk_Id_Contrato = c.Pk_Id_Contrato) AS Valor_Adiciones,
                c.Plazo AS Plazo_Inicial,
            (SELECT
                IFNULL(SUM(adiciones.Plazo), 0)
            FROM
                contratos_adiciones AS adiciones 
            WHERE
                adiciones.Fk_Id_Contrato = c.Pk_Id_Contrato) AS Plazo_Adiciones,
            c.Fecha_Inicial,
            c.Fecha_Vencimiento,
            c.Porcentaje_Avance,
            c.Acta_Inicio,
            c.Fecha_Acta_Inicio,
            t.Pk_Id_Terceros AS Pk_Id_Tercero_Contratante,
            t.Nombre AS Contratante
        FROM
            contratos AS c
        INNER JOIN tbl_terceros ON c.Fk_Id_Terceros = tbl_terceros.Pk_Id_Terceros
        INNER JOIN tbl_estados ON c.Fk_Id_Estado = tbl_estados.Pk_Id_Estado
        LEFT JOIN tbl_terceros AS t ON c.Fk_Id_Terceros_Contratante = t.Pk_Id_Terceros '.
        $id_contrato;        
        //Se retorna la consulta
        return $this->db->query($sql)->result(); 
    }//Fin listar_contratos
    
    /**
    * Lista la cantidad de p&oacute;lizas existentes.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_id_contratos($id_contratos){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->where('Pk_Id_Contrato', 1);
        
        //Se retorna la consulta
        $resultado = $this->db->get('contratos')->result();
        
        return $resultado;
    }//Fin listar_polizas()
    
    /**
    * Lista las p&oacute;lizas existentes.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_polizas(){
        //columnas que se van a retornar
        $this->db->select('*');
        
        //Se retorna la consulta
        $resultado = $this->db->get('poliza_tipo')->result();
        
        return $resultado;
    }//Fin listar_polizas()
    
    /**
    * Retorna los datos de la poliza de cumplimiento.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_poliza_cumplimiento($id_contrato){
        $sql = 
        'SELECT
            contratos_polizas.Pk_Id_Contratos_Poliza,
            contratos_polizas.Fk_Id_Poliza_Tipo,
            contratos_polizas.Numero,
            poliza_tipo.Poliza_Tipo,
            contratos_polizas.Fecha_Inicio,
            contratos_polizas.Fecha_Final,
            DATEDIFF(contratos_polizas.Fecha_Final, contratos_polizas.Fecha_Inicio) AS Vigencia,
            contratos_polizas.Valor,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre
        FROM
            contratos_polizas
        LEFT JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos_polizas.Fk_Id_Tercero
        WHERE
            contratos_polizas.Fk_Id_Contratos = '.$id_contrato.' AND
            contratos_polizas.Fk_Id_Poliza_Tipo = 1';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_poliza_cumplimiento()
    
    /**
    * Retorna los datos de la poliza de prestaciones sociales.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_poliza_prestaciones($id_contrato){
        $sql = 
        'SELECT
            contratos_polizas.Fk_Id_Poliza_Tipo,
            contratos_polizas.Numero,
            poliza_tipo.Poliza_Tipo,
            contratos_polizas.Fecha_Inicio,
            contratos_polizas.Fecha_Final,
            DATEDIFF(contratos_polizas.Fecha_Final, contratos_polizas.Fecha_Inicio) AS Vigencia,
            contratos_polizas.Valor,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre,
            contratos_polizas.Pk_Id_Contratos_Poliza
        FROM
            contratos_polizas
        LEFT JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos_polizas.Fk_Id_Tercero
        WHERE
            contratos_polizas.Fk_Id_Contratos = '.$id_contrato.' AND
            contratos_polizas.Fk_Id_Poliza_Tipo = 2';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_poliza_prestaciones()
    
    /**
    * Retorna los datos de la poliza de anticipos.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_poliza_anticipos($id_contrato){
        $sql = 
        'SELECT
            contratos_polizas.Fk_Id_Poliza_Tipo,
            contratos_polizas.Numero,
            poliza_tipo.Poliza_Tipo,
            contratos_polizas.Fecha_Inicio,
            contratos_polizas.Fecha_Final,
            DATEDIFF(contratos_polizas.Fecha_Final, contratos_polizas.Fecha_Inicio) AS Vigencia,
            contratos_polizas.Valor,
            contratos_polizas.Pk_Id_Contratos_Poliza,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre
        FROM
            contratos_polizas
        LEFT JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos_polizas.Fk_Id_Tercero
        WHERE
            contratos_polizas.Fk_Id_Contratos = '.$id_contrato.' AND
            contratos_polizas.Fk_Id_Poliza_Tipo = 3';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_poliza_anticipos()
    
    /**
    * Retorna los datos de la poliza de calidad.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_poliza_calidad($id_contrato){
        $sql = 
        'SELECT
            contratos_polizas.Fk_Id_Poliza_Tipo,
            contratos_polizas.Numero,
            poliza_tipo.Poliza_Tipo,
            contratos_polizas.Fecha_Inicio,
            contratos_polizas.Pk_Id_Contratos_Poliza,
            contratos_polizas.Fecha_Final,
            DATEDIFF(contratos_polizas.Fecha_Final, contratos_polizas.Fecha_Inicio) AS Vigencia,
            contratos_polizas.Valor,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre
        FROM
            contratos_polizas
        LEFT JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos_polizas.Fk_Id_Tercero
        WHERE
            contratos_polizas.Fk_Id_Contratos = '.$id_contrato.' AND
            contratos_polizas.Fk_Id_Poliza_Tipo = 4';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_poliza_calidad()
    
    /**
    * Retorna los datos de la poliza de estabilidad.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_poliza_estabilidad($id_contrato){
        $sql = 
        'SELECT
            contratos_polizas.Fk_Id_Poliza_Tipo,
            contratos_polizas.Numero,
            poliza_tipo.Poliza_Tipo,
            contratos_polizas.Fecha_Inicio,
            contratos_polizas.Pk_Id_Contratos_Poliza,
            contratos_polizas.Fecha_Final,
            DATEDIFF(contratos_polizas.Fecha_Final, contratos_polizas.Fecha_Inicio) AS Vigencia,
            contratos_polizas.Valor,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre
        FROM
            contratos_polizas
        LEFT JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos_polizas.Fk_Id_Tercero
        WHERE
            contratos_polizas.Fk_Id_Contratos = '.$id_contrato.' AND
            contratos_polizas.Fk_Id_Poliza_Tipo = 5';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_poliza_calidad()
    
    /**
    * Retorna los datos de la poliza de Responsabilidad Civil Contractual.
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_poliza_rc($id_contrato){
        $sql = 
        'SELECT
            contratos_polizas.Fk_Id_Poliza_Tipo,
            contratos_polizas.Numero,
            poliza_tipo.Poliza_Tipo,
            contratos_polizas.Fecha_Inicio,
            contratos_polizas.Pk_Id_Contratos_Poliza,
            contratos_polizas.Fecha_Final,
            DATEDIFF(contratos_polizas.Fecha_Final, contratos_polizas.Fecha_Inicio) AS Vigencia,
            contratos_polizas.Valor,
            tbl_terceros.Pk_Id_Terceros,
            tbl_terceros.Nombre
        FROM
            contratos_polizas
        LEFT JOIN poliza_tipo ON poliza_tipo.Pk_Id_Poliza_Tipo = contratos_polizas.Fk_Id_Poliza_Tipo
        LEFT JOIN tbl_terceros ON tbl_terceros.Pk_Id_Terceros = contratos_polizas.Fk_Id_Tercero
        WHERE
            contratos_polizas.Fk_Id_Contratos = '.$id_contrato.' AND
            contratos_polizas.Fk_Id_Poliza_Tipo = 6';
        
        //Se retorna la consulta
        return $this->db->query($sql)->result();
    }//Fin listar_poliza_rc()
    
    /**
    * Retorna los datos de las categorias de los contratos
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_categorias(){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->order_by('Categoria', 'asc');
        
        //Se retorna la consulta
        return $this->db->get('tbl_categorias')->result();
    }
    
    /**
    * Retorna los datos de las subcategorias de los contratos
    *
    * @access	
    * @param	
    * @param	
    * @return	
    */
    function listar_subcategorias($categoria){
        //columnas que se van a retornar
        $this->db->select('*');
        $this->db->where('Fk_Id_Categoria', $categoria);
        $this->db->order_by("Subcategoria", "asc");
        
        //Se retorna la consulta
        return $this->db->get('tbl_subcategorias')->result();
    }//Fin listar_subcategorias()

    /**
    * Calcula la fecha de vencimiento
    *
    * @access
    * @param    
    * @param    
    * @return   
    */
    function calcular_vencimiento($fecha_inicio, $dias){
        //Se realiza la consulta
        $sql =
        "SELECT DATE_ADD('{$fecha_inicio}', INTERVAL {$dias} DAY) AS Fecha_Vencimiento";

        //Se recorre
        foreach ($this->db->query($sql)->result() as $valor){
            //Se retorna la consulta
            return $valor->Fecha_Vencimiento;
        }
    }//Fin calcular_vencimiento

    /**
    * Guarda una adicion a un contrato especifico
    *
    * @access
    * @param    
    * @param    
    * @return   
    */
    function agregar_adicion($datos){
        $this->db->insert('contratos_adiciones', $datos);
    }//Fin agregar_adicion

    /**
    * Listas las adiciones de un contrato especifico
    *
    * @access
    * @param    
    * @param    
    * @return   
    */
    function listar_adiciones($id_contrato){
        $this->db->select('*');
        $this->db->where('Fk_Id_Contrato', $id_contrato);
        return $this->db->get('contratos_adiciones')->result();
    }//Fin listar_adiciones()
}//Fin contrato_model