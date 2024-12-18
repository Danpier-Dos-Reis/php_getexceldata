<?php
require 'Engine/Engine.php';

use Engine\Engine;

$filePath = __DIR__ . '/Test_Ventas.xlsx';
$engine = new Engine($filePath);
$sheetName = 'Hoja1'; // Nombre de la hoja en tu Excel

// Llamar a la función
$boletas = $engine->arrayBoletaVenta($filePath, $sheetName);

// echo $boletas;
// Mostrar los objetos creados
foreach ($boletas as $boleta) {
    echo "Id: {$boleta->Id}, Objeto: {$boleta->Objeto}, Cantidad: {$boleta->Cantidad}, Precio: {$boleta->Precio}";
    ?>
    <hr>
    <?php
}
?>