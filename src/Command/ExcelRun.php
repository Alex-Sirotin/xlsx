<?php
namespace ExcelCompare\Command;

use ExcelCompare\FileGenerator;
use ExcelCompare\Generators\SimpleXlsxGenerator;
use ExcelCompare\Generators\OneSheetGenerator;
use ExcelCompare\Generators\PhpSpreadsheetGenerator;
use ExcelCompare\Generators\EllumilelPhpExcelWriterGenerator;
use ExcelCompare\Generators\PhpXlsWriterGenerator;
use ExcelCompare\Generators\PeclXlsWriterGenerator;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Helper\Table;

class ExcelRun extends Command
{
    private $rowsCount = 1000;
    private $cellRepeat = 1;
    private $times = [];
    private $timestamp;
    private $uniqid;

    protected function configure()
    {
        $this->setName('excel:run')
            ->addOption('all', 'a', InputOption::VALUE_NONE, 'Generate all files')
            ->addOption(
                'simplexlsgen',
                's',
                InputOption::VALUE_NONE,
                'Generate by Shuchkin\SimpleXLSXGen'
            )
            ->addOption('onesheet', 'o', InputOption::VALUE_NONE, 'Generate by nimmneun\OneSheet')
            ->addOption(
                'ellumilelphpexcelwriter',
                'e',
                InputOption::VALUE_NONE,
                'Generate by ellumilel/php-excel-writer'
            )
            ->addOption('xlswriter', 'x', InputOption::VALUE_NONE, 'Generate by mk-j/PHP_XLSXWriter')
            ->addOption('peclxlswriter', 'p', InputOption::VALUE_NONE, 'Generate by Vtiful\Kernel\Excel')
            ->addOption('spreadsheet', 't', InputOption::VALUE_NONE, 'Generate by PHPOffice/PhpSpreadsheet')
            ->addOption('rows', 'r', InputOption::VALUE_REQUIRED, 'Rows count')
            ->addOption('cellrepeat', 'c', InputOption::VALUE_REQUIRED, 'Repeat cell');
        parent::configure();
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $rows = (int)$input->getOption('rows');
        if (is_numeric($rows) && $rows > 0) {
            $this->rowsCount = $rows;
        }
        $repeats = (int)$input->getOption('cellrepeat');
        if (is_numeric($repeats) && $repeats > 0) {
            $this->cellRepeat = $repeats;
        }
        $all = $input->getOption('all');

        $output->writeln('Generating...');
        $output->writeln("Rows: {$this->rowsCount}, Columns: {$this->cellRepeat}");
        $output->writeln('');
        $this->timestamp = time();
        $this->uniqid = uniqid();
        if ($input->getOption("simplexlsgen") || $all) {
            $this->runGenerator(new SimpleXlsxGenerator($output), 'SimpleXlsxGen.xlsx', $output);
        }
        if ($input->getOption("onesheet")) {
            $this->runGenerator(new OneSheetGenerator($output), 'OneSheet.xlsx', $output);
        }
        if ($input->getOption("ellumilelphpexcelwriter") || $all) {
            $this->runGenerator(
                new EllumilelPhpExcelWriterGenerator($output),
                'EllumilelPhpExcelWriterGenerator.xlsx',
                $output
            );
        }
        if ($input->getOption("xlswriter") || $all) {
            $this->runGenerator(new PhpXlsWriterGenerator($output), 'PhpXlsWriterGenerator.xlsx', $output);
        }
        if ($input->getOption("peclxlswriter") || $all) {
            $this->runGenerator(new PeclXlsWriterGenerator($output), 'PeclXlsWriterGenerator.xlsx', $output);
        }
        if ($input->getOption("spreadsheet") || $all) {
            $this->runGenerator(new PhpSpreadsheetGenerator($output), 'PhpSpreadsheetGenerator.xlsx', $output);
        }


        $this->printResult($output);
    }

    protected function printResult(OutputInterface $output)
    {
        $output->writeln('');
        $output->writeln('Result');

        $table = new Table($output);
        $header       = ['ID', 'Library', 'Time, ms', 'Rows', 'Columns'];
        $table->setHeaders($header);
        $i = 1;
        foreach ($this->times as $name => $value) {
            $table->addRow([
                $i++,
                $name,
                $value,
                $this->rowsCount,
                $this->cellRepeat
            ]);
        }

        $table->render();
        $output->writeln('');
    }

    protected function runGenerator(FileGenerator $generator, string $filename, OutputInterface $output)
    {
        $output->writeln($filename);
        $startTime = microtime(true);
        $fullName = "{$this->uniqid}_{$this->timestamp}_{$this->rowsCount}_{$this->cellRepeat}_{$filename}";
        $generator->setFilename($fullName);
        $generator->generate($this->rowsCount, $this->cellRepeat);
        $generator->save("output/{$fullName}");
        $this->times[$filename] = microtime(true) - $startTime;
        $output->writeln('');
        $output->writeln('');
    }
}
