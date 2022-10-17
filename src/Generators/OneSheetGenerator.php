<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use Symfony\Component\Console\Output\OutputInterface;

class OneSheetGenerator extends FileGenerator
{
    private $generator;
    private $row = [];

    public function __construct(OutputInterface $output)
    {
        parent::__construct($output);
        $this->generator = new \OneSheet\Writer();
        $this->generator->enableCellAutosizing();
        $this->row = [];
    }

    public function save(string $filename)
    {
        $this->generator->writeToFile($filename);
    }

    public function addRow(int $count = 1)
    {
        parent::addRow($count);
        $style = (new \OneSheet\Style\Style())
            ->setFontSize(13)
            ->setFontBold()
            ->setFontItalic()
            ->setFontUnderline()
            ->setFontColor('777777')
            ->setFillColor('CCCCFF');
        $this->generator->addRow($this->row, $style);
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
        $this->row[] = $this->data->getDate();
    }

    public function addTime()
    {
        $this->row[] = $this->data->getTime();
    }

    public function addDateTime()
    {
        $this->row[] = $this->data->getDateTime();
    }
}