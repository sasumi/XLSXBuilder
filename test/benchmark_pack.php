<?php

use LFPhp\XLSXBuilder\XLSXBuilder;
use function LFPhp\Func\dump;
use function LFPhp\Func\file_lines;
use function LFPhp\Func\format_size;
use function LFPhp\Func\read_csv_chunk;
use function LFPhp\Func\show_progress;

include "../vendor/autoload.php";

$csv_file = 'data_test.csv';
$output = "output_".time().".xlsx";
$line = file_lines($csv_file);

$start_time = time();
$mem = memory_get_usage(true);
echo "Start..., File Lines:$line ",date('H:i:s'),PHP_EOL;
echo "Process ID:", getmypid(),PHP_EOL;

XLSXBuilder::setTempDir(__DIR__.'/tmp');

$writer = new XLSXBuilder();
$sheet = $writer->createSheet('Sheet1');

XLSXBuilder::convertCSV($csv_file, $output);

echo "DONE,", date("H:i:s"),PHP_EOL;
echo "Time Cost:", time()-$start_time, 'sec', PHP_EOL;
echo "Mem Cost:", format_size(max($mem, memory_get_usage(true))), PHP_EOL;
echo "Output:", $output,PHP_EOL;

