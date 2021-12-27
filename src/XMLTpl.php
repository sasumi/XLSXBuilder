<?php
namespace LFPhp\XLSXBuilder;
use Exception;

/**
 * XML 模板渲染
 * @package LFPhp\XLSX
 */
class XMLTpl {
	public static function render($tpl_file, $params = []){
		if($params){
			extract($params);
		}
		ob_start();
		$tpl = __DIR__.'/tpl/'.$tpl_file;
		if(!is_file($tpl)){
			throw new Exception('XML template file no exists:'.$tpl);
		}
		include $tpl;
		return ob_get_clean();
	}
}