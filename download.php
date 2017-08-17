<?php
header("Content-Type:text/html;charset=utf-8");
require_once "function.php";
require_once "config.php";

download_xml(DOWNLOAD_DIR,iconv("UTF-8","GB2312",$_POST['xmlname']));
?>