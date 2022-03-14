<?php
use LFPhp\XLSXBuilder\Meta;
use function LFPhp\Func\xml_special_chars;
/** @var Meta $meta */
?>
<<?='?';?>xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF"><?=date("Y-m-d\TH:i:s.00\Z");?></dcterms:created>
<dc:title><?=xml_special_chars($meta->title);?></dc:title>
<dc:subject><?=xml_special_chars($meta->subject);?></dc:subject>
<dc:creator><?=xml_special_chars($meta->author);?></dc:creator>
<?php if ($meta->keywords):?>
	<cp:keywords><?=xml_special_chars(implode (", ",$meta->keywords));?></cp:keywords>
<?php endif;?>
<dc:description><?=xml_special_chars($meta->description);?></dc:description>
<cp:revision>0</cp:revision>
</cp:coreProperties>