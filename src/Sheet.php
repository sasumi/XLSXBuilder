<?php
namespace LFPhp\XLSXBuilder;

use function LFPhp\Func\xml_special_chars;

/**
 * @refs:
 * http://www.ecma-international.org/publications/standards/Ecma-376.htm
 * http://officeopenxml.com/SSstyles.php
 * http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
 */
class Sheet {
	const EXCEL_2007_MAX_ROW = 1048576;
	const EXCEL_2007_MAX_COL = 16384;
	const DEFAULT_SHEET_NAME = 'Sheet1';

	/** @var XLSXBuilder */
	private $xls_instance;

	public $file_name;
	public $sheet_name;
	public $xml_name;
	public $row_count;
	public $columns;
	public $merge_cells = [];
	private $max_cell_tag_start;
	private $max_cell_tag_end;
	private $auto_filter = false; //是否开启全局单元格过滤
	private $freeze_rows; //冻结开始行数
	private $freeze_columns; //冻结开始列数
	private $is_right_to_left_value = true; //数值靠右对齐

	//文件写入器
	private $file_writer;

	/**
	 * Sheet constructor.
	 * @param XLSXBuilder $xls_instance
	 * @param string $sheet_name 工作表名称
	 * @param number[] $col_widths 每列宽度设置
	 * @param bool $auto_filter 是否开启单元格数据过滤
	 * @param bool $freeze_rows 冻结开始行数
	 * @param bool $freeze_columns 冻结开始列数
	 * @param bool $is_right_to_left_value 数值靠右对齐
	 * @throws \Exception
	 */
	public function __construct(XLSXBuilder $xls_instance, $sheet_name, array $col_widths = [], $auto_filter = false, $freeze_rows = false, $freeze_columns = false, $is_right_to_left_value = false){
		$this->xls_instance = $xls_instance;
		$this->file_name = $xls_instance->createTempFile();
		$this->file_writer = new BufferWriter($this->file_name);
		$this->is_right_to_left_value = $is_right_to_left_value;
		$this->sheet_name = $sheet_name;
		$this->auto_filter = $auto_filter;
		$this->freeze_rows = $freeze_rows;
		$this->freeze_columns = $freeze_columns;
		$tab_selected = count($xls_instance->getSheets()) == 0;

		$str = '';
		$this->file_writer->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		$this->file_writer->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
		$this->file_writer->write('<sheetPr filterMode="false">');
		$this->file_writer->write('<pageSetUpPr fitToPage="false"/>');
		$this->file_writer->write('</sheetPr>');
		$this->max_cell_tag_start = $this->file_writer->ftell();
		$this->file_writer->write('<dimension ref="A1:'.Util::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL).'"/>');
		$this->max_cell_tag_end = $this->file_writer->ftell();

		$this->file_writer->write('<sheetViews>');
		$this->file_writer->write('<sheetView colorId = "64" defaultGridColor = "true" rightToLeft = "'.($is_right_to_left_value ? 'true':'false').'" showFormulas = "false" showGridLines = "true" showOutlineSymbols = "true" showRowColHeaders = "true" showZeros = "true" tabSelected = "'.($tab_selected ? 'true': 'false').'" topLeftCell = "A1" view = "normal" windowProtection = "false" workbookViewId = "0" zoomScale = "100" zoomScaleNormal = "100" zoomScalePageLayoutView = "100" >');
		if($this->freeze_rows && $this->freeze_columns){
			$this->file_writer->write('<pane ySplit="'.$this->freeze_rows.'" xSplit="'.$this->freeze_columns.'" topLeftCell="'.Util::xlsCell($this->freeze_rows, $this->freeze_columns).'" activePane="bottomRight" state="frozen"/>');
			$this->file_writer->write('<selection activeCell = "'.Util::xlsCell($this->freeze_rows, 0).'" activeCellId = "0" pane = "topRight" sqref = "'.Util::xlsCell($this->freeze_rows, 0).'"/>');
			$this->file_writer->write('<selection activeCell = "'.Util::xlsCell(0, $this->freeze_columns).'" activeCellId = "0" pane = "bottomLeft" sqref = "'.Util::xlsCell(0, $this->freeze_columns).'"/>');
			$this->file_writer->write('<selection activeCell = "'.Util::xlsCell($this->freeze_rows, $this->freeze_columns).'" activeCellId = "0" pane = "bottomRight" sqref = "'.Util::xlsCell($this->freeze_rows, $this->freeze_columns).'"/>');
		}elseif($this->freeze_rows){
			$this->file_writer->write('<pane ySplit="'.$this->freeze_rows.'" topLeftCell="'.Util::xlsCell($this->freeze_rows, 0).'" activePane="bottomLeft" state="frozen"/>');
			$this->file_writer->write('<selection activeCell = "'.Util::xlsCell($this->freeze_rows, 0).'" activeCellId = "0" pane = "bottomLeft" sqref = "'.Util::xlsCell($this->freeze_rows, 0).'"/>');
		}elseif($this->freeze_columns){
			$this->file_writer->write('<pane xSplit = "'.$this->freeze_columns.'" topLeftCell = "'.Util::xlsCell(0, $this->freeze_columns).'" activePane = "topRight" state = "frozen"/>');
			$this->file_writer->write('<selection activeCell = "'.Util::xlsCell(0, $this->freeze_columns).'" activeCellId = "0" pane = "topRight" sqref = "'.Util::xlsCell(0, $this->freeze_columns).'"/>');
		}else{ // not frozen
			$this->file_writer->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
		}
		$this->file_writer->write('</sheetView>');
		$this->file_writer->write('</sheetViews>');
		$this->file_writer->write('<cols>');
		$i = 0;
		if(!empty($col_widths)){
			foreach($col_widths as $column_width){
				$this->file_writer->write('<col collapsed="false" hidden="false" max="'.($i + 1).'" min="'.($i + 1).'" style="0" customWidth="true" width="'.floatval($column_width).'"/>');
				$i++;
			}
		}
		$this->file_writer->write('<col collapsed = "false" hidden = "false" max = "1024" min = "'.($i + 1).'" style = "0" customWidth = "false" width = "11.5"/>');
		$this->file_writer->write('</cols>');
		$this->file_writer->write(' <sheetData>');
		$this->file_writer->write($str);
	}

	/**
	 * @param $str
	 */
	public function writeString($str){
		$this->file_writer->write($str);
	}

	/**
	 * 最终写入文件
	 */
	public function finalize(){
		$this->file_writer->write('</sheetData>');
		if(!empty($this->merge_cells)){
			$this->file_writer->write('<mergeCells>');
			foreach($this->merge_cells as $range){
				$this->file_writer->write('<mergeCell ref="'.$range.'"/>');
			}
			$this->file_writer->write('</mergeCells>');
		}

		$max_cell = Util::xlsCell($this->row_count - 1, count($this->columns) - 1);

		if($this->auto_filter){
			$this->file_writer->write('<autoFilter ref="A1:'.$max_cell.'"/>');
		}

		$this->file_writer->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
		$this->file_writer->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
		$this->file_writer->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
		$this->file_writer->write('<headerFooter differentFirst="false" differentOddEven="false">');
		$this->file_writer->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
		$this->file_writer->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
		$this->file_writer->write('</headerFooter>');
		$this->file_writer->write('</worksheet>');

		$max_cell_tag = '<dimension ref="A1:'.$max_cell.'"/>';
		$padding_length = $this->max_cell_tag_end - $this->max_cell_tag_start - strlen($max_cell_tag);
		$this->file_writer->fseek($this->max_cell_tag_start);

		$this->file_writer->write($max_cell_tag.str_repeat(" ", $padding_length));
		$this->file_writer->close();
	}

	/**
	 * 单元格写入数据
	 * @param int $row_number
	 * @param int $column_number
	 * @param mixed $value
	 * @param string $num_format_type
	 * @param $cell_style_idx
	 */
	public function writeCell($row_number, $column_number, $value, $num_format_type, $cell_style_idx){
		$cell_name = Util::xlsCell($row_number, $column_number);
		if(!is_scalar($value) || $value === ''){ //objects, array, empty
			$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'"/>');
		}elseif(is_string($value) && $value[0] == '='){
			$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="s"><f>'.xml_special_chars($value).'</f></c>');
		}elseif($num_format_type == 'n_date'){
			$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.intval(Util::convertDateTime($value)).'</v></c>');
		}elseif($num_format_type == 'n_datetime'){
			$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.Util::convertDateTime($value).'</v></c>');
		}elseif($num_format_type == 'n_numeric'){
			$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.xml_special_chars($value).'</v></c>');//int,float,currency
		}elseif($num_format_type == 'n_string'){
			$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="inlineStr"><is><t>'.xml_special_chars($value).'</t></is></c>');
		}elseif($num_format_type == 'n_auto' || 1){ //auto-detect unknown column types
			if(!is_string($value) || $value == '0' || ($value[0] != '0' && ctype_digit($value)) || preg_match("/^-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)){
				$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.xml_special_chars($value).'</v></c>');//int,float,currency
			}else{ //implied: ($cell_format=='string')
				$this->file_writer->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="inlineStr"><is><t>'.xml_special_chars($value).'</t></is></c>');
			}
		}
	}

	/**
	 * 写入行数据
	 * @param array $row
	 * @param array $row_options
	 */
	public function writeRow(array $row, array $row_options = []){
		if (count($this->columns) < count($row)) {
			$default_column_types = $this->initializeColumnTypes( array_fill($from=0, $until=count($row), 'GENERAL') );//will map to n_auto
			$this->columns = array_merge((array)$this->columns, $default_column_types);
		}

		if ($row_options){
			$ht = isset($row_options['height']) ? floatval($row_options['height']) : 12.1;
			$customHt = isset($row_options['height']);
			$hidden = isset($row_options['hidden']) ? (bool)($row_options['hidden']) : false;
			$collapsed = isset($row_options['collapsed']) ? (bool)($row_options['collapsed']) : false;
			$this->file_writer->write('<row collapsed="'.($collapsed ? 'true' : 'false').'" customFormat="false" customHeight="'.($customHt ? 'true' : 'false').'" hidden="'.($hidden ? 'true' : 'false').'" ht="'.($ht).'" outlineLevel="0" r="'.($this->row_count + 1).'">');
		}else{
			$this->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.($this->row_count + 1).'">');
		}
		$c=0;
		foreach ($row as $v) {
			$number_format = $this->columns[$c]['number_format'];
			$number_format_type = $this->columns[$c]['number_format_type'];
			$cell_style_idx = empty($row_options) ? $this->columns[$c]['default_cell_style'] : $this->xls_instance->addCellStyle( $number_format, json_encode(isset($row_options[0]) ? $row_options[$c] : $row_options) );
			$this->writeCell($this->row_count, $c, $v, $number_format_type, $cell_style_idx);
			$c++;
		}
		$this->file_writer->write('</row>');
		$this->row_count++;
	}


	/**
	 * 初始化栏信息
	 * @param $header_types
	 * @return array
	 */
	private function initializeColumnTypes($header_types){
		$column_types = [];
		foreach($header_types as $v){
			$number_format = Util::numberFormatStandardized($v);
			$number_format_type = Util::determineNumberFormatType($number_format);
			$cell_style_idx = $this->xls_instance->addCellStyle($number_format, $style_string = null);
			$column_types[] = array(
				'number_format'      => $number_format,//contains excel format like 'YYYY-MM-DD HH:MM:SS'
				'number_format_type' => $number_format_type, //contains friendly format like 'datetime'
				'default_cell_style' => $cell_style_idx,
			);
		}
		return $column_types;
	}

	/**
	 * 标记单元格合并信息
	 * @param $start_cell_row
	 * @param $start_cell_column
	 * @param $end_cell_row
	 * @param $end_cell_column
	 */
	public function markMergedCell($start_cell_row, $start_cell_column, $end_cell_row, $end_cell_column){
		$startCell = Util::xlsCell($start_cell_row, $start_cell_column);
		$endCell = Util::xlsCell($end_cell_row, $end_cell_column);
		$this->merge_cells[] = $startCell.":".$endCell;
	}


	/**
	 * 写入工作表头部信息
	 * @param array $header_types
	 * @param null $col_options
	 * @throws \Exception
	 */
	public function setHeader(array $header_types, $col_options = null){
		$suppress_row = isset($col_options['suppress_row']) ? intval($col_options['suppress_row']) : false;
		if(is_bool($col_options)){
			$suppress_row = intval($col_options);
		}
		$style = &$col_options;
		$this->columns = $this->initializeColumnTypes($header_types);
		if(!$suppress_row){
			$header_row = array_keys($header_types);
			$this->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
			foreach ($header_row as $c => $v) {
				$cell_style_idx = empty($style) ? $this->columns[$c]['default_cell_style'] : $this->xls_instance->addCellStyle( 'GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style) );
				$this->writeCell(0, $c, $v, $number_format_type='n_string', $cell_style_idx);
			}
			$this->file_writer->write('</row>');
			$this->row_count++;
		}
	}

	public function getFileName(){
		return $this->file_name;
	}

	/**
	 * @return bool
	 */
	public function isAutoFilter(){
		return $this->auto_filter;
	}
}