# XLSX 生成库
> 用于生成XLSX（ Open XML standard）格式文件。尤其在大量数据需要保存为Excel可读取文件时，可使用该生成库生成。可有效避免内存占用过大问题。

## 1. 安装

1. PHP 版本大于或等于 5.6
2. 必须安装扩展：mb_string、php_json、php_zip

请使用Composer进行安装：
```shell script
composer require lfphp/xlsxbuilder
```

## 2. 使用

注意：该库在使用过程依赖临时目录，用于存储生成过程的临时文件。

若需要指定该临时目录，请使用 `$writer->setTempDir()` 进行设置，缺省情况使用系统temp目录。

代码库使用方法：

```php
//设置列格式
$header_types = array(
   'created'=>'date',
   'product_id'=>'string',
   'quantity'=>'#,##0',
   'amount'=>'price',
   'description'=>'string',
   'tax'=>'[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
);
//数据
$data = array(
   array('2015-01-01',874234242342343,1,'44.00','misc','=D2*0.05'),
   array('2015-01-12',324,2,'88.00','none','=D3*0.05'),
);

$writer = new XLSXBuilder();
$sheet = $writer->createSheet('Sheet1'); //创建工作表 Sheet1
$sheet->setHeader($header_types);
foreach($data as $row){
   $writer->getSheet('Sheet1')->writeRow($row);
}
$writer->saveAs('example.xlsx');
```

## 3. 性能测试

| 测试项目    | 数量      | 耗时 | CPU峰值 | 内存占用（峰值） | 环境             |
| ----------- | --------- | ---- | ------- | ---------------- | ---------------- |
| csv文件转换 | 1,000,000 | 57秒 | 14%     | 4MB              | win10+PHP7.3.5TS |