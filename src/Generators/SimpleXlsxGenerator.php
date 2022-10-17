<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use Shuchkin\SimpleXLSXGen;

class SimpleXlsxGenerator extends FileGenerator
{
    private $generator;
    private $rows = [];
    private $row = [];

    public function __construct()
    {
        parent::__construct();
        $this->generator = new SimpleXLSXGen();
        $this->rows = [];
        $this->row = [];
    }

    function save(string $filename)
    {
        $this->generator->addSheet($this->rows, 'SimpleXLSXGen');
        $this->generator->saveAs($filename);
    }

    function addRow(int $count = 1)
    {
        parent::addRow($count);
        $this->rows[] = $this->row;
        $this->row = [];
    }

    function addString()
    {
        $this->row[] = $this->data->getText();
    }

    function addNumber()
    {
        $this->row[] = $this->data->getNumber();
    }

    function addFloat()
    {
        $this->row[] = $this->data->getFloat();
    }

    function addDate()
    {
        $this->row[] = $this->data->getDate();
    }

    function addTime()
    {
        $this->row[] = $this->data->getTime();
    }

    function addDateTime()
    {
        $this->row[] = $this->data->getDateTime();
    }
}