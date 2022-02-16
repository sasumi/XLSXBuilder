<?php
namespace LFPhp\XLSXBuilder;
use Exception;
use ZipArchive;
use function LFPhp\Func\read_csv_chunk;
use function LFPhp\Func\xml_special_chars;

class XLSXBuilder {
	/** @var Meta $meta */
	public $meta;

	/** @var Sheet[] */
	protected $sheets = [];
	protected $temp_files = [];
	protected $temp_dir = null;
	protected $cell_styles = [];
	protected $number_formats = [];

	public function __construct(Meta $meta = null){
		defined('ENT_XML1') or define('ENT_XML1', 16);//for php 5.3, avoid fatal error
		date_default_timezone_get() or date_default_timezone_set('UTC');//php.ini missing tz, avoid warning
		$this->meta = $meta ?: new Meta();
		$this->setTempDir();
		$this->addCellStyle($number_format = 'GENERAL', $style_string = null);
	}

	/**
	 * 析构函数，删除所有临时文件
	 */
	public function __destruct(){
		foreach($this->temp_files as $temp_file){
			@unlink($temp_file);
		}
	}

	/**
	 * 创建临时文件
	 * @return string 临时文件名
	 * @throws \Exception
	 */
	public function createTempFile(){
		$temp_dir = !empty($this->temp_dir) ? $this->temp_dir : sys_get_temp_dir();
		$filename = tempnam($temp_dir, "xlsx_writer_");
		if(!$filename){
			throw new Exception("Unable to create temp file in $temp_dir");
		}
		$this->temp_files[] = $filename;
		return $filename;
	}

	/**
	 * 转换CSV文件到xlsx
	 * @param string $csv_file CSV文件名称
	 * @param string $xlsx_filename 保存的xlsx文件全名（包含路径）
	 * @param array $sheet_header 工作表栏目信息，格式如下：
		* $header = array(
		* 'created'=>'date',
		* 'product_id'=>'integer',
		* 'quantity'=>'#,##0',
		* 'amount'=>'price',
		* 'description'=>'string',
		* 'tax'=>'[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
		* );
	 * @param int $ignore_csv_head_lines 忽略CSV文件中的文件头部行数
	 * @throws \Exception
	 */
	public static function convertCSV($csv_file, $xlsx_filename, $sheet_header = [], $ignore_csv_head_lines = 0){
		$builder = new static();
		$sheet = $builder->createSheet('Sheet1');
		if($sheet_header){
			$sheet->setHeader($sheet_header);
		}
		read_csv_chunk(function($rows)use($sheet){
			foreach($rows as $row){
				$sheet->writeRow($row);
			}
		}, $csv_file, [], 1000, $ignore_csv_head_lines);
		$builder->saveAs($xlsx_filename);
	}

	/**
	 * 保存所有文件
	 * @return string xlxs 文件路径
	 * @throws \Exception
	 */
	public function save(){
		$temp_file = $this->createTempFile();
		self::saveAs($temp_file);
		return $temp_file;
	}

	/**
	 * 写入到指定文件
	 * @param $filename
	 * @throws \Exception
	 */
	public function saveAs($filename){
		if(!$this->sheets){
			throw new Exception("no worksheets defined.");
		}
		if(file_exists($filename)){
			if(is_writable($filename)){
				@unlink($filename); //if the zip already exists, remove it
			}else{
				throw new Exception("file is not writeable.");
			}
		}
		foreach($this->sheets as $sheet) {
			$sheet->finalize();
		}
		$zip = new ZipArchive();
		if(!$zip->open($filename, ZipArchive::CREATE)){
			throw new Exception("unable to create zip.");
		}

		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml", XMLTpl::render("docProps/app.xml.php", ['meta'=>$this->meta]));
		$zip->addFromString("docProps/core.xml", XMLTpl::render('docProps/core.xml.php', ['meta'=>$this->meta]));

		$zip->addEmptyDir("_rels/");
		$zip->addFromString("_rels/.rels", XMLTpl::render('_rels/.rels.php'));

		$zip->addEmptyDir("xl/worksheets/");
		foreach($this->sheets as $sheet){
			$zip->addFile($sheet->getFileName(), "xl/worksheets/".$sheet->xml_name);
		}
		$zip->addFromString("xl/workbook.xml", self::buildWorkbookXML());
		$zip->addFile($this->writeStylesXML(), "xl/styles.xml");
		$zip->addFromString("[Content_Types].xml", XMLTpl::render('[Content_Types].xml.php', ['sheets'=>$this->sheets]));

		$zip->addEmptyDir("xl/_rels/");
		$zip->addFromString("xl/_rels/workbook.xml.rels", XMLTpl::render('xl/_rels/workbook.xml.rels.php', ['sheets'=>$this->sheets] ));
		$zip->close();
	}

	/**
	 * 初始化工作表
	 * @param $sheet_name
	 * @param array $col_widths
	 * @param bool $auto_filter
	 * @param bool $freeze_rows
	 * @param bool $freeze_columns
	 * @return Sheet
	 * @throws Exception
	 */
	public function createSheet($sheet_name = Sheet::DEFAULT_SHEET_NAME, $col_widths = [], $auto_filter = false, $freeze_rows = false, $freeze_columns = false){
		$sheet_xml_name = 'sheet' . (count($this->sheets) + 1).".xml";
		$sheet = new Sheet($this, $sheet_name, $col_widths, $auto_filter, $freeze_rows, $freeze_columns);
		$sheet->xml_name = $sheet_xml_name;
		$sheet->row_count = 0;
		$sheet->columns = [];
		$sheet->merge_cells = [];
		$this->sheets[] = $sheet;
		return $sheet;
	}

	/**
	 * 获取指定表名的工作表
	 * @param $sheet_name
	 * @return Sheet
	 */
	public function getSheet($sheet_name){
		return $this->sheets[$sheet_name];
	}

	/**
	 * 添加单元格样式
	 * @param $number_format
	 * @param $cell_style_string
	 * @return false|int|string
	 */
	public function addCellStyle($number_format, $cell_style_string){
		$number_format_idx = Util::addToListGetIndex($this->number_formats, $number_format);
		$lookup_string = $number_format_idx.";".$cell_style_string;
		return Util::addToListGetIndex($this->cell_styles, $lookup_string);
	}

	/**
	 * 生成字体样式索引
	 * @return array
	 */
	protected function styleFontIndexes(){
		static $border_allowed = array('left','right','top','bottom');
		static $border_style_allowed = array('thin','medium','thick','dashDot','dashDotDot','dashed','dotted','double','hair','mediumDashDot','mediumDashDotDot','mediumDashed','slantDashDot');
		static $horizontal_allowed = array('general','left','right','justify','center');
		static $vertical_allowed = array('bottom','center','distributed','top');
		$default_font = array('size'=>'10','name'=>'Arial','family'=>'2');
		$fills = array('','');//2 placeholders for static xml later
		$fonts = array('','','','');//4 placeholders for static xml later
		$borders = array('');//1 placeholder for static xml later
		$style_indexes = [];
		foreach($this->cell_styles as $i => $cell_style_string){
			$semi_colon_pos = strpos($cell_style_string, ";");
			$number_format_idx = substr($cell_style_string, 0, $semi_colon_pos);
			$style_json_string = substr($cell_style_string, $semi_colon_pos + 1);
			$style = @json_decode($style_json_string, $as_assoc = true);

			$style_indexes[$i] = array('num_fmt_idx' => $number_format_idx);//initialize entry
			if(isset($style['border']) && is_string($style['border']))//border is a comma delimited str
			{
				$border_value['side'] = array_intersect(explode(",", $style['border']), $border_allowed);
				if(isset($style['border-style']) && in_array($style['border-style'], $border_style_allowed)){
					$border_value['style'] = $style['border-style'];
				}
				if(isset($style['border-color']) && is_string($style['border-color']) && $style['border-color'][0] == '#'){
					$v = substr($style['border-color'], 1, 6);
					$v = strlen($v) == 3 ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v;// expand cf0 => ccff00
					$border_value['color'] = "FF".strtoupper($v);
				}
				$style_indexes[$i]['border_idx'] = Util::addToListGetIndex($borders, json_encode($border_value));
			}
			if(isset($style['fill']) && is_string($style['fill']) && $style['fill'][0] == '#'){
				$v = substr($style['fill'], 1, 6);
				$v = strlen($v) == 3 ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v;// expand cf0 => ccff00
				$style_indexes[$i]['fill_idx'] = Util::addToListGetIndex($fills, "FF".strtoupper($v));
			}
			if(isset($style['halign']) && in_array($style['halign'], $horizontal_allowed)){
				$style_indexes[$i]['alignment'] = true;
				$style_indexes[$i]['halign'] = $style['halign'];
			}
			if(isset($style['valign']) && in_array($style['valign'], $vertical_allowed)){
				$style_indexes[$i]['alignment'] = true;
				$style_indexes[$i]['valign'] = $style['valign'];
			}
			if(isset($style['wrap_text'])){
				$style_indexes[$i]['alignment'] = true;
				$style_indexes[$i]['wrap_text'] = (bool)$style['wrap_text'];
			}

			$font = $default_font;
			if(isset($style['font-size'])){
				$font['size'] = floatval($style['font-size']);//floatval to allow "10.5" etc
			}
			if(isset($style['font']) && is_string($style['font'])){
				if($style['font'] == 'Comic Sans MS'){
					$font['family'] = 4;
				}
				if($style['font'] == 'Times New Roman'){
					$font['family'] = 1;
				}
				if($style['font'] == 'Courier New'){
					$font['family'] = 3;
				}
				$font['name'] = strval($style['font']);
			}
			if(isset($style['font-style']) && is_string($style['font-style'])){
				if(strpos($style['font-style'], 'bold') !== false){
					$font['bold'] = true;
				}
				if(strpos($style['font-style'], 'italic') !== false){
					$font['italic'] = true;
				}
				if(strpos($style['font-style'], 'strike') !== false){
					$font['strike'] = true;
				}
				if(strpos($style['font-style'], 'underline') !== false){
					$font['underline'] = true;
				}
			}
			if(isset($style['color']) && is_string($style['color']) && $style['color'][0] == '#'){
				$v = substr($style['color'], 1, 6);
				$v = strlen($v) == 3 ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v;// expand cf0 => ccff00
				$font['color'] = "FF".strtoupper($v);
			}
			if($font != $default_font){
				$style_indexes[$i]['font_idx'] = Util::addToListGetIndex($fonts, json_encode($font));
			}
		}
		return array('fills'=>$fills,'fonts'=>$fonts,'borders'=>$borders,'styles'=>$style_indexes );
	}

	/**
	 * 写入样式到XML文件
	 * @return string
	 * @throws \Exception
	 */
	protected function writeStylesXML(){
		$r = $this->styleFontIndexes();
		$fills = $r['fills'];
		$fonts = $r['fonts'];
		$borders = $r['borders'];
		$style_indexes = $r['styles'];

		$temporary_filename = $this->createTempFile();
		$file = new BufferWriter($temporary_filename);
		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		$file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		$file->write('<numFmts count="'.count($this->number_formats).'">');
		foreach($this->number_formats as $i=>$v) {
			$file->write('<numFmt numFmtId="'.(164+$i).'" formatCode="'.xml_special_chars($v).'" />');
		}
		$file->write('</numFmts>');

		$file->write('<fonts count="'.(count($fonts)).'">');
		$file->write(		'<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');

		foreach($fonts as $font) {
			if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
				$f = json_decode($font,true);
				$file->write('<font>');
				$file->write(	'<name val="'.htmlspecialchars($f['name']).'"/><charset val="1"/><family val="'.intval($f['family']).'"/>');
				$file->write(	'<sz val="'.intval($f['size']).'"/>');
				if (!empty($f['color'])) { $file->write('<color rgb="'.strval($f['color']).'"/>'); }
				if (!empty($f['bold'])) { $file->write('<b val="true"/>'); }
				if (!empty($f['italic'])) { $file->write('<i val="true"/>'); }
				if (!empty($f['underline'])) { $file->write('<u val="single"/>'); }
				if (!empty($f['strike'])) { $file->write('<strike val="true"/>'); }
				$file->write('</font>');
			}
		}
		$file->write('</fonts>');

		$file->write('<fills count="'.(count($fills)).'">');
		$file->write(	'<fill><patternFill patternType="none"/></fill>');
		$file->write(	'<fill><patternFill patternType="gray125"/></fill>');
		foreach($fills as $fill) {
			if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
				$file->write('<fill><patternFill patternType="solid"><fgColor rgb="'.strval($fill).'"/><bgColor indexed="64"/></patternFill></fill>');
			}
		}
		$file->write('</fills>');

		$file->write('<borders count="'.(count($borders)).'">');
		$file->write(    '<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>');
		foreach($borders as $border) {
			if (!empty($border)) { //fonts have an empty placeholder in the array to offset the static xml entry above
				$pieces = json_decode($border,true);
				$border_style = !empty($pieces['style']) ? $pieces['style'] : 'hair';
				$border_color = !empty($pieces['color']) ? '<color rgb="'.strval($pieces['color']).'"/>' : '';
				$file->write('<border diagonalDown="false" diagonalUp="false">');
				foreach (array('left', 'right', 'top', 'bottom') as $side)
				{
					$show_side = in_array($side,$pieces['side']);
					$file->write($show_side ? "<$side style=\"$border_style\">$border_color</$side>" : "<$side/>");
				}
				$file->write(  '<diagonal/>');
				$file->write('</border>');
			}
		}
		$file->write('</borders>');

		$file->write('<cellStyleXfs count="20">');
		$file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
		$file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
		$file->write('<protection hidden="false" locked="true"/>');
		$file->write('</xf>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
		$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
		$file->write('</cellStyleXfs>');

		$file->write('<cellXfs count="'.(count($style_indexes)).'">');
		foreach($style_indexes as $v){
			$applyAlignment = isset($v['alignment']) ? 'true' : 'false';
			$wrapText = !empty($v['wrap_text']) ? 'true' : 'false';
			$horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
			$vertAlignment = isset($v['valign']) ? $v['valign'] : 'bottom';
			$applyBorder = isset($v['border_idx']) ? 'true' : 'false';
			$applyFont = 'true';
			$borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
			$fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
			$fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
			$file->write('<xf applyAlignment="'.$applyAlignment.'" applyBorder="'.$applyBorder.'" applyFont="'.$applyFont.'" applyProtection="false" borderId="'.($borderIdx).'" fillId="'.($fillIdx).'" fontId="'.($fontIdx).'" numFmtId="'.(164+$v['num_fmt_idx']).'" xfId="0">');
			$file->write('	<alignment horizontal="'.$horizAlignment.'" vertical="'.$vertAlignment.'" textRotation="0" wrapText="'.$wrapText.'" indent="0" shrinkToFit="false"/>');
			$file->write('	<protection locked="true" hidden="false"/>');
			$file->write('</xf>');
		}
		$file->write('</cellXfs>');
		$file->write(	'<cellStyles count="6">');
		$file->write(		'<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
		$file->write(		'<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
		$file->write(		'<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
		$file->write(		'<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
		$file->write(		'<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
		$file->write(		'<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
		$file->write(	'</cellStyles>');
		$file->write('</styleSheet>');
		$file->close();
		return $temporary_filename;
	}

	protected function buildWorkbookXML(){
		$i = 0;
		$workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$workbook_xml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		$workbook_xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
		$workbook_xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
		$workbook_xml .= '<sheets>';
		foreach($this->sheets as $sheet_name=>$sheet) {
			$sheet_name = Util::sanitizeSheetName($sheet->sheet_name);
			$workbook_xml.='<sheet name="'.xml_special_chars($sheet_name).'" sheetId="'.($i+1).'" state="visible" r:id="rId'.($i+2).'"/>';
			$i++;
		}
		$workbook_xml.='</sheets>';
		$workbook_xml.='<definedNames>';
		foreach($this->sheets as $sheet_name=>$sheet) {
			if ($sheet->isAutoFilter()) {
				$sheet_name = Util::sanitizeSheetName($sheet->sheet_name);
				$workbook_xml.='<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''.xml_special_chars($sheet_name).'\'!$A$1:' . Util::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1, true) . '</definedName>';
				$i++;
			}
		}
		$workbook_xml.='</definedNames>';
		$workbook_xml.='<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
		return $workbook_xml;
	}

	/**
	 * 设置临时存储目录
	 * @param string|null $temp_dir
	 * @throws \Exception
	 */
	public function setTempDir($temp_dir = null){
		if($temp_dir && !is_dir($temp_dir)){
			mkdir($temp_dir, 0777, true);
		}
		$temp_dir = $temp_dir ?: sys_get_temp_dir();
		if(!is_writable($temp_dir)){
			throw new Exception("Temporary directory is no writeable:$temp_dir");
		}
		$this->temp_dir = $temp_dir;
	}

	/**
	 * 获取所有工作表
	 * @return \LFPhp\XLSXBuilder\Sheet[]
	 */
	public function getSheets(){
		return $this->sheets;
	}
}