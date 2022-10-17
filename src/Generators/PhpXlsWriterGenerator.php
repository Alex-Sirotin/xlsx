<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use Symfony\Component\Console\Output\OutputInterface;
use XLSXWriter;

class PhpXlsWriterGenerator extends FileGenerator
{
    private $generator;
    private $row = [];
    private $headers = [];
    private $headersAdded = false;

    public function __construct(OutputInterface $output)
    {
        parent::__construct($output);
        $this->generator = new XLSXWriter();
        $this->row = [];
        $this->headers = [];
    }

    public function save(string $filename = null)
    {
        $this->generator->writeToFile($filename);
    }

    public function addRow(int $count = 1)
    {
        parent::addRow($count);
        if (!$this->headersAdded) {
            $this->headersAdded = true;
            $this->generator->writeSheetHeader('XLSXWriter', $this->headers);
        }
        $this->generator->writeSheetRow('XLSXWriter', $this->row);
        $this->row = [];
    }

    public function addString()
    {
        $this->row[] = $this->data->getText();
        $this->headers[] = 'string';
    }

    public function addNumber()
    {
        $this->row[] = $this->data->getNumber();
        $this->headers[] = 'integer';
    }

    public function addFloat()
    {
        $this->row[] = $this->data->getFloat();
        $this->headers[] = 'price';
    }

    public function addDate()
    {
        $this->row[] = $this->data->getDateTime()->format('c');
        $this->headers[] = 'DD.MM.YYYY';
    }

    public function addTime()
    {
        $this->row[] = $this->data->getDateTime()->format('c');
        $this->headers[] = 'HH:MM';
    }

    public function addDateTime()
    {
        $this->row[] = $this->data->getDateTime()->format('c');
        $this->headers[] = 'DD.MM.YYYY (HH:MM:SS)';
    }
}