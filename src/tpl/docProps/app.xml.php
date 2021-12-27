<?php
use LFPhp\XLSXBuilder\Meta;
use function LFPhp\Func\xml_special_chars;
/** @var Meta $meta */
?>
<<?='?';?>xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<TotalTime>0</TotalTime>
<Company><?echo xml_special_chars($meta->company);?></Company>
</Properties>