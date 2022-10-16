<?php

require __DIR__.'/vendor/autoload.php';

use Symfony\Component\Console\Application;
use ExcelCompare\Command\ExcelRun;

//use Monolog\Level;
//use Monolog\Logger;
//use Monolog\Handler\StreamHandler;

//// create a log channel
//$log = new Logger('name');
//$log->pushHandler(new StreamHandler('path/to/your.log', Level::Warning));
//
//// add records to the log
//$log->warning('Foo');
//$log->error('Bar');

$app = new Application();

$app->add(new ExcelRun());

try {
    $app->run();
} catch (\Exception $ex) {

}


