<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use Shuchkin\SimpleXLSXGen;
use Symfony\Component\Console\Output\OutputInterface;

class SimpleXlsxGenerator extends FileGenerator
{
    private $generator;
    private $rows = [];
    private $row = [];

    public function __construct(OutputInterface $output)
    {
        parent::__construct($output);
        $this->generator = new SimpleXLSXGen();
        $this->rows = [];
        $this->row = [];
    }

    public function save(string $filename = null)
    {
        $this->generator->addSheet($this->rows, 'SimpleXLSXGen');
        $this->generator->saveAs($filename);
    }

    public function addRow(int $count = 1)
    {
        parent::addRow($count);
        $this->rows[] = $this->row;
        $this->row = [];
    }

    public function addString()
    {
        $this->row[] = $this->data->getText();
    }

    public function addNumber()
    {
        $this->row[] = $this->data->getNumber();
    }

    public function addFloat()
    {
        $this->row[] = $this->data->getFloat();
    }

    public function addDate()
    {
        $this->row[] = $this->data->getDateTime()->format('Y-m-d');
    }

    public function addTime()
    {
        $this->row[] = $this->data->getDateTime()->format('H:i:s');
    }

    public function addDateTime()
    {
        $this->row[] = $this->data->getDateTime()->format('Y-m-d H:i:s');
    }
}