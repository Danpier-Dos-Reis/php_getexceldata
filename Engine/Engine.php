<?php
namespace Engine;

require 'Models/BoletaVenta.php';
require 'vendor/autoload.php'; // Asegúrate de cargar PhpSpreadsheet con Composer

use PhpOffice\PhpSpreadsheet\IOFactory;
use Models\BoletaVenta;

class Engine {
    public $columns = []; // Array para guardar nombres de las columnas

    public function __construct($filePath) {
        $this->loadExcel($filePath);
    }

    private function loadExcel($filePath) {
        try {
            // Cargar el archivo Excel
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getActiveSheet();

            // Obtener la primera fila (encabezados)
            $firstRow = $sheet->getRowIterator(1, 1)->current();
            $cellIterator = $firstRow->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false); //incluye celdas vacias entre el rago de columnas

            // Guardar los valores de la primera fila en $columns
            foreach ($cellIterator as $cell) {
                $this->columns[] = $cell->getValue();
            }

        } catch (\Exception $e) {
            die("Error al procesar el archivo: " . $e->getMessage());
        }
    }

    public function getColumns() {
        return $this->columns;
    }

    function getNumbOfSheets($filePath) {
        try {
            // Cargar el archivo Excel
            $spreadsheet = IOFactory::load($filePath);
    
            // Obtener la cantidad de hojas
            $sheetCount = $spreadsheet->getSheetCount();
    
            return $sheetCount;
    
        } catch (\Exception $e) {
            // En caso de error, muestra un mensaje
            die("Error al procesar el archivo: " . $e->getMessage());
        }
    }

    function getNumbOfColumns($filePath) {
        try {
            // Cargar el archivo Excel
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getActiveSheet();
    
            // Obtener la cantidad de columnas (última columna con datos)
            $highestColumn = $sheet->getHighestColumn(); // Retorna la letra de la última columna
            $columnCount = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
    
            return $columnCount;
    
        } catch (\Exception $e) {
            die("Error al procesar el archivo: " . $e->getMessage());
        }
    }

    function getNumbOfRows($filePath) {
        try {
            // Cargar el archivo Excel
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getActiveSheet();
    
            // Obtener la cantidad de filas (última fila con datos)
            $rowCount = $sheet->getHighestRow();
    
            return $rowCount;
    
        } catch (\Exception $e) {
            die("Error al procesar el archivo: " . $e->getMessage());
        }
    }

    function ReadSheetTargeted($filePath, $sheetName) {
        try {
            // Cargar el archivo Excel
            $spreadsheet = IOFactory::load($filePath);
            
            // Obtener la hoja por su nombre
            $sheet = $spreadsheet->getSheetByName($sheetName);
            if (!$sheet) {
                throw new Exception("No se encontró la hoja: $sheetName");
            }
    
            // Leer los datos de la hoja
            foreach ($sheet->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false);
    
                foreach ($cellIterator as $cell) {
                    echo $cell->getValue() . "\t";
                }
                echo "\n";
            }
        } catch (\Exception $e) {
            die("Error al leer la hoja: " . $e->getMessage());
        }
    }

    /**
    * Crea un array de objetos BoletaVenta a partir de una hoja de Excel.
    *
    * @param string $filePath Ruta del archivo Excel.
    * @param string $sheetName Nombre de la hoja específica en el Excel.
    * @return array Array de objetos BoletaVenta.
    * @throws Exception Si el archivo no se puede procesar o la hoja no existe.
    */
    function arrayBoletaVenta($filePath, $sheetName) {
        $boletas = []; // Array para almacenar objetos BoletaVenta
    
        try {
            // Cargar el archivo Excel
            $spreadsheet = IOFactory::load($filePath);
    
            // Obtener la hoja específica
            $sheet = $spreadsheet->getSheetByName($sheetName);
            if (!$sheet) {
                throw new Exception("No se encontró la hoja: $sheetName");
            }
    
            // Obtener el número máximo de filas y empezar desde la fila 2 (omitiendo encabezado)
            $highestRow = $sheet->getHighestRow();
    
            for ($row = 2; $row <= $highestRow; $row++) {
                // Leer valores de las celdas de la fila
                $id       = $sheet->getCell("A$row")->getValue();
                $objeto   = $sheet->getCell("B$row")->getValue();
                $cantidad = $sheet->getCell("C$row")->getValue();
                $precio   = $sheet->getCell("D$row")->getValue();
    
                // Crear una instancia de BoletaVenta
                $boleta = new BoletaVenta($id, $objeto, $cantidad, $precio);
    
                // Agregar al array
                $boletas[] = $boleta;
            }
        } catch (\Exception $e) {
            die("Error al procesar el archivo: " . $e->getMessage());
        }
    
        return $boletas;
    }
}