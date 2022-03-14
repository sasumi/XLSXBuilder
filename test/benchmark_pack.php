<?php

use LFPhp\XLSXBuilder\XLSXBuilder;
use function LFPhp\Func\dump;
use function LFPhp\Func\file_lines;
use function LFPhp\Func\format_size;
use function LFPhp\Func\read_csv_chunk;
use function LFPhp\Func\show_progress;

include "../vendor/autoload.php";
dump(format_size(memory_get_usage()));

XLSXBuilder::setTempDir(__DIR__.'/tmp');

$writer = new XLSXBuilder();
$sheet = $writer->createSheet('Sheet1');

dump(format_size(memory_get_usage()));
$csv_file = 'data_100.csv';
$output = "output_".time().".xlsx";
XLSXBuilder::convertCSV($csv_file, $output);

echo $output,PHP_EOL;
die('DONE');