<?php

use LFPhp\XLSXBuilder\XLSXBuilder;
include '../vendor/autoload.php';

$header_types = array(
	'created'=>'date',
	'product_id'=>'string',
	'quantity'=>'#,##0',
	'amount'=>'price',
	'description'=>'string',
	'tax'=>'[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
);

$data = array(
	array('2015-01-01',874234242342343,1,'44.00','misc','=D2*0.05'),
	array('2015-01-12',324,2,'88.00','none','=D3*0.05'),
);

$writer = new XLSXBuilder();
$sheet = $writer->createSheet('Sheet1');
$sheet->setHeader($header_types);
foreach($data as $row){
	$writer->getSheet('Sheet1')->writeRow($row);
}
$writer->saveAs('example.xlsx');