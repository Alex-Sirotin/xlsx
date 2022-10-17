<?php
namespace ExcelCompare;

use ExcelCompare\ExcelData;
use Symfony\Component\Console\Helper\ProgressBar;
use Symfony\Component\Console\Output\OutputInterface;

abstract class FileGenerator
{
    /** @var ExcelData */
    protected $data;
    /** @var OutputInterface */
    private $output;
    /** @var string */
    protected $filename;

    public function __construct(OutputInterface $output)
    {
        $this->data = new ExcelData();
        $this->output = $output;
    }

    public function setFilename($filename)
    {
        $this->filename = $filename;
    }

    public function addRow(int $count = 1)
    {
        foreach (range(1, $count) as $i) {
            $this->addNumber();
        }
        foreach (range(1, $count) as $i) {
            $this->addFloat();
        }
        foreach (range(1, $count) as $i) {
            $this->addDate();
        }
        foreach (range(1, $count) as $i) {
            $this->addTime();
        }
        foreach (range(1, $count) as $i) {
            $this->addDateTime();
        }
        foreach (range(1, $count) as $i) {
            $this->addString();
        }
    }

    public function generate(int $rowCount, int $repeat = 1)
    {
        $progress = new ProgressBar($this->output, $rowCount);
        $progress->start();
        foreach (range(1, $rowCount) as $i) {
            $this->addRow($repeat);
            $progress->advance();
        }
        $progress->finish();
    }

    abstract public function save(string $filename = null);
    abstract public function addString();
    abstract public function addNumber();
    abstract public function addFloat();
    abstract public function addDate();
    abstract public function addTime();
    abstract public function addDateTime();
}
