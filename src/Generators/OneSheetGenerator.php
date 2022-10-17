<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;

class OneSheetGenerator extends FileGenerator
{
    private $generator;
    private $row = [];

    public function __construct()
    {
        parent::__construct();
        $this->generator = new \OneSheet\Writer();
        $this->generator->enableCellAutosizing();
        $this->row = [];
    }

    function save(string $filename)
    {
        $this->generator->writeToFile($filename);
    }

    function addRow(int $count = 1)
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