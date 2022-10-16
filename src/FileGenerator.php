<?php
namespace ExcelCompare;

use ExcelCompare\ExcelData;

abstract class FileGenerator
{
//    private $headers = [];
//    private $rows = [];
    /** @var ExcelData */
    private $data;

    public function __construct(int $rowCount)
    {
        $data = new ExcelData($rowCount);
//        $this->headers = $headers;
//        $this->rows = $rows;
    }

    private function fill()
    {
//        $this->
    }

    abstract function generate();
}