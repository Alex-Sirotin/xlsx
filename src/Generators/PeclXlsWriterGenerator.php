<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use Symfony\Component\Console\Output\OutputInterface;

class PeclXlsWriterGenerator extends FileGenerator
{
    private $generator;
    private $excel;
    private $row = 0;
    private $col = 0;

    public function __construct(OutputInterface $output)
    {
        parent::__construct($output);
        $config = [
            'path' => './output'
        ];

        $this->excel = new \Vtiful\Kernel\Excel($config);
    }

    public function setFilename($filename)
    {
        parent::setFilename($filename);
        $this->generator = $this->excel->fileName($this->filename, 'XlsWriter(PECL)');
    }

    public function save(string $filename = null)
    {
        $this->generator->output();
    }

    public function addRow(int $count = 1)
    {
        parent::addRow($count);
        $this->col = 0;
        $this->row++;
    }

    public function addString()
    {
        $this->generator->insertText($this->row, $this->col++, $this->data->getText());
    }

    public function addNumber()
    {
        $this->generator->insertText($this->row, $this->col++, $this->data->getNumber(), "#,##0");
    }

    public function addFloat()
    {
        $this->generator->insertText($this->row, $this->col++, $this->data->getFloat(), "#,##0.00");
    }

    public function addDate()
    {
        $this->generator->insertDate($this->row, $this->col++, $this->data->getDateTime()->getTimestamp(), "dd.mm.YYyy");
    }

    public function addTime()
    {
        $this->generator->insertDate($this->row, $this->col++, $this->data->getDateTime()->getTimestamp(), 'hh:mm:ss');
    }

    public function addDateTime()
    {
        $this->generator->insertDate($this->row, $this->col++, $this->data->getDateTime()->getTimestamp(), 'ddd, dd.mmm.yy (hh:mm:ss)');
    }
}
