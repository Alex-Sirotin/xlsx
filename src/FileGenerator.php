<?php
namespace ExcelCompare;

use ExcelCompare\ExcelData;

abstract class FileGenerator
{
    /** @var ExcelData */
    protected $data;

    public function __construct()
    {
        $this->data = new ExcelData();
    }

    protected function addRow(int $count = 1)
    {
        foreach(range(1, $count) as $i) $this->addNumber();
        foreach(range(1, $count) as $i) $this->addFloat();
        foreach(range(1, $count) as $i) $this->addDate();
        foreach(range(1, $count) as $i) $this->addTime();
        foreach(range(1, $count) as $i) $this->addDateTime();
        foreach(range(1, $count) as $i) $this->addString();
    }

    public function generate(int $rowCount, int $repeat = 1)
    {
        foreach(range(1, $rowCount) as $i) {
            $this->addRow($repeat);
        }
    }

    abstract function save(string $filename);
//    abstract function addRow(int $id, int $repeat);

    abstract function addString();
    abstract function addNumber();
    abstract function addFloat();
    abstract function addDate();
    abstract function addTime();
    abstract function addDateTime();
//    abstract function addLeft($value);
//    abstract function addRight($value);
//    abstract function addCenter($value);
}