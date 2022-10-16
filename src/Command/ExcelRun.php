<?php
namespace ExcelCompare\Command;

use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Helper\ProgressBar;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;

class ExcelRun extends Command
{
    protected function configure()
    {
        $this->setName('excel:run');
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $output->writeln('Success!');
    }
}