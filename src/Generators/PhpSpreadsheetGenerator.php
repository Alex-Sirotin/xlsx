<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\Console\Output\OutputInterface;

class PhpSpreadsheetGenerator extends FileGenerator
{
    private $generator;
    private $rowIndex = 1;
    private $colIndex = 1;
    private $spreadsheet;

    public function __construct(OutputInterface $output)
    {
        parent::__construct($output);

        //\PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder(new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder());
        $this->spreadsheet = new Spreadsheet();
        $this->generator = $this->spreadsheet->getActiveSheet();
    }

    public function save(string $filename = null)
    {
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($filename);
    }

    public function addRow(int $count = 1)
    {
        parent::addRow($count);
        $this->rowIndex++;
        $this->colIndex = 1;
    }

    public function addString()
    {
        $this->generator->setCellValueByColumnAndRow($this->colIndex++, $this->rowIndex, $this->data->getText());
    }

    public function addNumber()
    {
        $this->generator->setCellValueByColumnAndRow($this->colIndex, $this->rowIndex, $this->data->getNumber());
        $this->generator
            ->getCellByColumnAndRow($this->colIndex++, $this->rowIndex)
            ->getStyle()
            ->getNumberFormat()
            ->setFormatCode(
                \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER
            );
    }

    public function addFloat()
    {
        $this->generator->setCellValueByColumnAndRow($this->colIndex, $this->rowIndex, $this->data->getFloat());
        $this->generator
            ->getCellByColumnAndRow($this->colIndex++, $this->rowIndex)
            ->getStyle()
            ->getNumberFormat()
            ->setFormatCode(
                \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_00
            );
    }

    public function addDate()
    {
        $this->generator->setCellValueByColumnAndRow(
            $this->colIndex,
            $this->rowIndex,
            \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($this->data->getDateTime())
        );
        $this->generator
            ->getCellByColumnAndRow($this->colIndex++, $this->rowIndex)
            ->getStyle()
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DMYMINUS);
    }

    public function addTime()
    {
        $this->generator->setCellValueByColumnAndRow(
            $this->colIndex,
            $this->rowIndex,
            \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($this->data->getDateTime())
        );
        $this->generator
            ->getCellByColumnAndRow($this->colIndex++, $this->rowIndex)
            ->getStyle()
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_TIME6);
    }

    public function addDateTime()
    {
        $this->generator->setCellValueByColumnAndRow(
            $this->colIndex,
            $this->rowIndex,
            \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($this->data->getDateTime())
        );
        $this->generator
            ->getCellByColumnAndRow($this->colIndex++, $this->rowIndex)
            ->getStyle()
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DATETIME);
    }
}
