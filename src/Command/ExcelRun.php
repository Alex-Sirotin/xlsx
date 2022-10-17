<?php
namespace ExcelCompare\Command;

use \ExcelCompare\Generators\SimpleXlsxGenerator;
use \ExcelCompare\Generators\OneSheetGenerator;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;

class ExcelRun extends Command
{
    private $rowsCount = 1000;

    protected function configure()
    {
        $this->setName('excel:run')
            ->addOption('simplexlsgen', 's', InputOption::VALUE_NONE, 'Generate by Shuchkin\SimpleXLSXGen')
            ->addOption('onesheet', 'o', InputOption::VALUE_NONE, 'Generate by nimmneun\OneSheet')
            ->addOption('rows', 'r', InputOption::VALUE_REQUIRED, 'Rows count');
        parent::configure();
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $rows = (int)$input->getOption('rows');
        if (is_numeric($rows) && $rows > 0) {
            $this->rowsCount = $rows;
        }
        if ($input->getOption("simplexlsgen")) {
            $this->runSimpleXlsGen();
        }
        if ($input->getOption("onesheet")) {
            $this->runOneSheet();
        }
    }

    protected function runSimpleXlsGen()
    {
        $generator = new SimpleXlsxGenerator();
        $generator->generate($this->rowsCount);
        $generator->save('output/SimpleXlsxGen.xlsx');
    }

    protected function runOneSheet()
    {
        $generator = new OneSheetGenerator();
        $generator->generate($this->rowsCount);
        $generator->save('output/OneSheet.xlsx');
    }
}