<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use Ellumilel\ExcelWriter;

class EllumilelPhpExcelWriterGenerator extends FileGenerator
{
    private $generator;
    private $row = [];
    private $headers = [];
    private $headersAdded = false;

    public function __construct()
    {
        parent::__construct();
        $this->generator = new ExcelWriter();
        $this->row = [];
        $this->headers = [];
    }

    function save(string $filename)
    {
        $this->generator->writeToFile($filename);
    }

    function addRow(int $count = 1)
    {
        parent::addRow($count);
        if (!$this->headersAdded) {
            $this->headersAdded = true;
            $this->generator->writeSheetHeader('Ellumilel', $this->headers);
        }
        $this->generator->writeSheetRow('Ellumilel', $this->row);
        $this->row = [];
    }

    function addString()
    {
        $this->row[] = $this->data->getText();
        $this->headers[] = 'string';
    }

    function addNumber()
    {
        $this->row[] = $this->data->getNumber();
        $this->headers[] = 'integer';
    }

    function addFloat()
    {
        $this->row[] = $this->data->getFloat();
        $this->headers[] = 'float';
    }

    function addDate()
    {
        $this->row[] = $this->data->getDate();
        $this->headers[] = 'DD.MM.YYYY';
    }

    function addTime()
    {
        $this->row[] = $this->data->getTime();
        $this->headers[] = 'string';
    }

    function addDateTime()
    {
        $this->row[] = $this->data->getDateTime()->format('c');
        $this->headers[] = 'DD.MM.YYYY (HH::MM:SS)';
    }
}