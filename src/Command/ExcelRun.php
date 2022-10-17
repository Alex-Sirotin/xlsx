<?php
namespace ExcelCompare\Command;

use ExcelCompare\FileGenerator;
use ExcelCompare\Generators\SimpleXlsxGenerator;
use ExcelCompare\Generators\OneSheetGenerator;
use ExcelCompare\Generators\EllumilelPhpExcelWriterGenerator;
use ExcelCompare\Generators\PhpXlsWriterGenerator;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Symfony\Component\Console\Helper\Table;

class ExcelRun extends Command
{
    private $rowsCount = 1000;
    private $times = [];

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
            ->addOption('rows', 'r', InputOption::VALUE_REQUIRED, 'Rows count');
        parent::configure();
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $rows = (int)$input->getOption('rows');
        if (is_numeric($rows) && $rows > 0) {
            $this->rowsCount = $rows;
        }
        $all = $input->getOption('all');

        $output->writeln('Generating...');
        $output->writeln('');
        if ($input->getOption("simplexlsgen") || $all) {
            $this->runGenerator(new SimpleXlsxGenerator($output), 'output/SimpleXlsxGen.xlsx', $output);
        }
        if ($input->getOption("onesheet") || $all) {
            $this->runGenerator(new OneSheetGenerator($output), 'output/OneSheet.xlsx', $output);
        }
        if ($input->getOption("ellumilelphpexcelwriter") || $all) {
            $this->runGenerator(
                new EllumilelPhpExcelWriterGenerator($output),
                'output/EllumilelPhpExcelWriterGenerator.xlsx',
                $output
            );
        }
        if ($input->getOption("xlswriter") || $all) {
            $this->runGenerator(new PhpXlsWriterGenerator($output), 'output/PhpXlsWriterGenerator.xlsx', $output);
        }
        $this->printResult($output);
    }

    protected function printResult(OutputInterface $output)
    {
        $output->writeln('');
        $output->writeln('Result');

        $table = new Table($output);
        $header       = ['ID', 'Library', 'Time, ms', 'Rows'];
        $table->setHeaders($header);
        $i = 1;
        foreach ($this->times as $name => $value) {
            $table->addRow([
                $i++,
                $name,
                $value,
                $this->rowsCount
            ]);
        }

        $table->render();
        $output->writeln('');
    }

    protected function runGenerator(FileGenerator $generator, string $filename, OutputInterface $output)
    {
        $output->writeln($filename);
        $startTime = microtime(true);
        $generator->generate($this->rowsCount);
        $generator->save($filename);
        $this->times[$filename] = microtime(true) - $startTime;
        $output->writeln('');
        $output->writeln('');
    }
}
