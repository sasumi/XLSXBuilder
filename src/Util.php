<?php
namespace LFPhp\XLSXBuilder;
class Util {
	/**
	 * 清理工作表名称
	 * @param $sheet_name
	 * @return string
	 */
	public static function sanitizeSheetName($sheet_name){
		static $bad_chars = '\\/?*:[]';
		static $good_chars = '        ';
		$sheet_name = strtr($sheet_name, $bad_chars, $good_chars);
		$sheet_name = function_exists('mb_substr') ? mb_substr($sheet_name, 0, 31) : substr($sheet_name, 0, 31);
		$sheet_name = trim(trim(trim($sheet_name), "'"));//trim before and after trimming single quotes
		return !empty($sheet_name) ? $sheet_name : 'Sheet'.((rand()%900) + 100);
	}

	public static function determineNumberFormatType($num_format){
		$num_format = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)]/i", "", $num_format);
		if($num_format == 'GENERAL')
			return 'n_auto';
		if($num_format == '@')
			return 'n_string';
		if($num_format == '0')
			return 'n_numeric';
		if(preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $num_format))
			return 'n_datetime';
		if(preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $num_format))
			return 'n_datetime';
		if(preg_match('/[Y]{2,4}(?![^"]*+")/i', $num_format))
			return 'n_date';
		if(preg_match('/[D]{1,2}(?![^"]*+")/i', $num_format))
			return 'n_date';
		if(preg_match('/[M]{1,2}(?![^"]*+")/i', $num_format))
			return 'n_date';
		if(preg_match('/$(?![^"]*+")/', $num_format))
			return 'n_numeric';
		if(preg_match('/%(?![^"]*+")/', $num_format))
			return 'n_numeric';
		if(preg_match('/0(?![^"]*+")/', $num_format))
			return 'n_numeric';
		return 'n_auto';
	}

	/**
	 * 单元格数据格式标准化
	 * @param $num_format
	 * @return string
	 */
	public static function numberFormatStandardized($num_format){
		$format_str_map = [
			'number'   => '0',
			'integer'  => '0',
			'string'   => '@',
			'date'     => 'YYYY-MM-DD',
			'datetime' => 'YYYY-MM-DD HH:MM:SS',
			'time'     => 'HH:MM:SS',
			'price'    => '#,##0.00',
			'money'    => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
			'dollar'   => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
			'euro'     => '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]',
		];
		$fmt = isset($format_str_map[$num_format]) ? $format_str_map[$num_format] : $num_format;
		$ignore_until = '';
		$escaped = '';
		for($i = 0, $ix = strlen($fmt); $i < $ix; $i++){
			$c = $fmt[$i];
			if($ignore_until == '' && $c == '[')
				$ignore_until = ']';
			else if($ignore_until == '' && $c == '"')
				$ignore_until = '"';
			else if($ignore_until == $c)
				$ignore_until = '';
			if($ignore_until == '' && ($c == ' ' || $c == '-' || $c == '(' || $c == ')') && ($i == 0 || $num_format[$i - 1] != '_'))
				$escaped .= "\\".$c;
			else
				$escaped .= $c;
		}
		return $escaped;
	}

	/**
	 * @param $row_number int, zero based
	 * @param $column_number int, zero based
	 * @param $absolute bool
	 * @return string Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
	 * */
	public static function xlsCell($row_number, $column_number, $absolute = false){
		$n = $column_number;
		for($r = ""; $n >= 0; $n = intval($n/26) - 1){
			$r = chr($n%26 + 0x41).$r;
		}
		if($absolute){
			return '$'.$r.'$'.($row_number + 1);
		}
		return $r.($row_number + 1);
	}

	/**
	 * @param $haystack
	 * @param $needle
	 * @return false|int|string
	 */
	public static function addToListGetIndex(&$haystack, $needle){
		$existing_idx = array_search($needle, $haystack, $strict = true);
		if($existing_idx === false){
			$existing_idx = count($haystack);
			$haystack[] = $needle;
		}
		return $existing_idx;
	}

	/**
	 * @param $date_input
	 * @return float|int|mixed
	 * thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
	 */
	public static function convertDateTime($date_input){
		$seconds = 0;    # Time expressed as fraction of 24h hours in seconds
		$year = $month = $day = 0;

		$date_time = $date_input;
		if(preg_match("/(\d{4})-(\d{2})-(\d{2})/", $date_time, $matches)){
			list($junk, $year, $month, $day) = $matches;
		}
		if(preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches)){
			list($junk, $hour, $min, $sec) = $matches;
			$seconds = ($hour*60*60 + $min*60 + $sec)/(24*60*60);
		}

		//using 1900 as epoch, not 1904, ignoring 1904 special case

		# Special cases for Excel.
		if("$year-$month-$day" == '1899-12-31')
			return $seconds;    # Excel 1900 epoch
		if("$year-$month-$day" == '1900-01-00')
			return $seconds;    # Excel 1900 epoch
		if("$year-$month-$day" == '1900-02-29')
			return 60 + $seconds;    # Excel false leapday

		# We calculate the date by calculating the number of days since the epoch
		# and adjust for the number of leap days. We calculate the number of leap
		# days by normalising the year in relation to the epoch. Thus the year 2000
		# becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
		$epoch = 1900;
		$offset = 0;
		$norm = 300;
		$range = $year - $epoch;

		# Set month days and check for leap year.
		$leap = (($year%400 == 0) || (($year%4 == 0) && ($year%100))) ? 1 : 0;
		$mon_days_map = array(31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

		# Some boundary checks
		if($year != 0 || $month != 0 || $day != 0){
			if($year < $epoch || $year > 9999)
				return 0;
			if($month < 1 || $month > 12)
				return 0;
			if($day < 1 || $day > $mon_days_map[$month - 1])
				return 0;
		}

		# Accumulate the number of days since the epoch.
		$days = $day;    # Add days for current month
		$days += array_sum(array_slice($mon_days_map, 0, $month - 1));    # Add days for past months
		$days += $range*365;                      # Add days for past years
		$days += intval(($range)/4);             # Add leapdays
		$days -= intval(($range + $offset)/100); # Subtract 100 year leapdays
		$days += intval(($range + $offset + $norm)/400);  # Add 400 year leapdays
		$days -= $leap;                                      # Already counted above

		# Adjust for Excel erroneously treating 1900 as a leap year.
		if($days > 59){
			$days++;
		}

		return $days + $seconds;
	}
}