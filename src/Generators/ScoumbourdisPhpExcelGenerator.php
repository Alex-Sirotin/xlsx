<?php

namespace ExcelCompare\Generators;

use ExcelCompare\FileGenerator;
use PHPExcel;
use PHPExcel_Writer_Excel2007;
use Symfony\Component\Console\Output\OutputInterface;

class ScoumbourdisPhpExcelGenerator extends FileGenerator
{
    private $generator;
    private $xls;
    private $rowIndex = 1;
    private $colIndex = 0;

    public function __construct(OutputInterface $output)
    {
        parent::__construct($output);

        $this->xls = new PHPExcel();
        $this->generator = $this->xls->getActiveSheet();
    }

    public function save(string $filename = null)
    {
        $objWriter = new PHPExcel_Writer_Excel2007($this->xls);
        $objWriter->save($filename);
    }

    public function addRow(int $count = 1)
    {
        parent::addRow($count);
        $this->rowIndex++;
        $this->colIndex = 0;
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
