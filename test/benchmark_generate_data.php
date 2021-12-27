<?php

use function LFPhp\Func\rand_string;
use function LFPhp\Func\show_progress;

include "../vendor/autoload.php";

$buff = '';
$tmp = getopt('n:');
$total = $tmp['n'];
if(!$total){
	die('Example: php '.$_SERVER['SCRIPT_NAME'].' -n1000');
}
$i = $total;
$sm = ['待激活', '正常', '注销'];
$data_file = "data_{$total}.csv";
$fp = fopen($data_file, 'w');
while($i-- > 0){
	$buff .= "$i,u_".rand_string(32).",手机尾号".rand(0,$i)."用户,+853-".rand_string(11,'1234567890').",".$sm[array_rand($sm, 1)].",B端导入,\n";
	if($i % 10000){
		fwrite($fp, $buff);
		$buff = '';
	}
	show_progress($total-$i, $total);
}
if($buff){
	fwrite($fp, $buff);
}
fclose($fp);
echo "DONE";
exit;