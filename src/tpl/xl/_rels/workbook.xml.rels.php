<?php
/** @var Sheet[] $sheets */
use LFPhp\XLSXBuilder\Sheet;
?>
<<?='?';?>xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<?php  $i = 0; foreach($sheets as $sheet):?>
<Relationship Id="rId<?=($i + 2);?>" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/<?=$sheet->xml_name;?>"/>
<?php $i++; endforeach;?>
</Relationships>