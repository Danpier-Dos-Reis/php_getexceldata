<?php
namespace Models;

class BoletaVenta {
    public $Id;
    public $Objeto;
    public $Cantidad;
    public $Precio;

    public function __construct($id, $objeto, $cantidad, $precio) {
        $this->Id = $id;
        $this->Objeto = $objeto;
        $this->Cantidad = $cantidad;
        $this->Precio = $precio;
    }
}