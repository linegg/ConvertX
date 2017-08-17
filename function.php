<?php 
require_once 'config.php';
header("Content-Type:text/html;charset=UTF-8");
function abend()
{
	echo "<h3>转换格式失败！！！</h3>";
	echo "<p>请在<a href='./helpDoc/help.html'>注意事项</a>查看你的文件是否符合规范！</p>";
}

//开放5个自定义字段
function get_cus($arrpost)
{
	$cus = array();
	if(!empty($arrpost["cus_1"]))
	{
		$cus[1] = $arrpost["cus_1"];
	}
	if(!empty($arrpost["cus_2"]))
	{
		$cus[2] = $arrpost["cus_2"];
	}
	if(!empty($arrpost["cus_3"]))
	{
		$cus[3] = $arrpost["cus_3"];
	}
	if(!empty($arrpost["cus_4"]))
	{
		$cus[4] = $arrpost["cus_4"];
	}
	if(!empty($arrpost["cus_5"]))
	{
		$cus[5] = $arrpost["cus_5"];
	}
	return $cus;
}
//写入到xml
function template_xml($tcinfo,$fp,$cus)
{
	$n = count($cus);
	$act = array();
	$exp = array();
	$pre = array();
	
	$temp = "\n<testcase internalid=\" \" name=\"{$tcinfo[TC_NAME_C]}\">
	<node_order><![CDATA[]]></node_order>
	<externalid><![CDATA[]]></externalid>
	<version><![CDATA[]]></version>
	<summary><![CDATA[<p>{$tcinfo[SUMMARY_C]}</p>]]></summary>
	<preconditions><![CDATA[";
	/*给preg_split增加limit参数，避免“1、xxx1、”情况下内容丢失问题--20150312*/
	$pre = preg_split("/1、/",$tcinfo[PRECONDITION_C],2);
	$s = 2;
	if($pre[0] != "")
	{
		$temp = $temp."<p>".$pre[0]."</p>";
	}
	while($pre[1] != "")
	{
		$pre = preg_split("/{$s}、/",$pre[1],2);
		$k = $s - 1;
		$temp = $temp."<p>".$k."、".$pre[0]."</p>";
		$s++;
	}
	$temp = $temp."]]></preconditions>
	<execution_type><![CDATA[{$tcinfo[EXECUTION_TYPE_C]}]]></execution_type>
	<importance><![CDATA[{$tcinfo[PRIORITY_C]}]]></importance>
	<estimated_exec_duration>{$tcinfo[ESTIMATED_TIME_C]}</estimated_exec_duration>
	<status>1</status>
	<steps>
		<step>
			<step_number><![CDATA[1]]></step_number>
			<actions><![CDATA[";
	$act = preg_split("/1、/",$tcinfo[ACTION_C],2);
	$s = 2;
	if($act[0] != "")
	{
		$temp = $temp."<p>".$act[0]."</p>";
	}
	while($act[1] != "")
	{
		$act = preg_split("/{$s}、/",$act[1],2);
		$k = $s - 1;
		$temp = $temp."<p>".$k."、".$act[0]."</p>";
		$s++;
	}
	
	$temp = $temp."]]></actions>
			<expectedresults><![CDATA[";
			
	$exp = preg_split("/1、/",$tcinfo[EXPECTED_RESULTS_C],2);
	$s = 2;
	if($exp[0] != "")
	{
		$temp = $temp."<p>".$exp[0]."</p>";
	}
	while($exp[1] != "")
	{
		$exp = preg_split("/{$s}、/",$exp[1],2);
		$k = $s - 1;
		$temp = $temp."<p>".$k."、".$exp[0]."</p>";
		$s++;
	}
	/*
	$exp = 	preg_split("/[1-9]、/",$tcinfo[EXPECTED_RESULTS_C]);
	for($i = 0;$i <= count($exp) - 1;$i ++)
	{
		if($i == 0)
		{
			$temp = $temp."<p>".$exp[$i]."</p>";
		}
		else
		{
			$temp = $temp."<p>".$i."、".$exp[$i]."</p>";
		}
	}
	*/
	$temp = $temp."]]></expectedresults>
			<execution_type><![CDATA[1]]></execution_type>
		</step>
	</steps>
	";
	$temp = $temp."<custom_fields>";
	if($n)
	{
		for($i = 1;$i <= $n;$i++)
		{
			$key = "CUS_".$i;
			$temp = $temp."<custom_field>
			<name><![CDATA[{$cus[$i]}]]></name>
			<value><![CDATA[{$tcinfo[$key]}]]></value>
			</custom_field>
			";
		}
	}
	$temp = $temp."<custom_field>
	<name><![CDATA[".AUTHOR."]]></name>
	<value><![CDATA[{$tcinfo[AUTHOR_C]}]]></value>
	</custom_field></custom_fields>
	";
	
	$require = explode(";",$tcinfo[REQUIREMENT_C]);
	$n = count($require);
	$temp = $temp."<requirements>";
	for($i = 1;$i <= $n;$i++)
	{
		$temp = $temp."<requirement>
		<doc_id><![CDATA[{$require[$i - 1]}]]></doc_id>
		</requirement>";	
	}
	$temp =$temp."</requirements>";
	$temp = $temp."</testcase>";
	fwrite($fp,$temp);
}
//写入<testsuite>到xml文件
function template_xml2($suitename,$fp,$source,$level)
{
	$temp = "
	<testsuite name=\"{$suitename}\">
	<node_order><![CDATA[]]></node_order>
	<details><![CDATA[]]></details>
	";
	if($source)
	{
		$temp = "</testsuite>".$temp;
		if($level == 1 && $source == 2)
		{
			$temp = "</testsuite>".$temp;
			$source --;
		}
		$source --;
	}
	fwrite($fp,$temp);
	$source ++;
	return $source;
}
//根据<testsuite>个数写入</testsuite>
function template_end($fp,$source)
{
	for($i = 1;$i <= $source + 1;$i ++)
	{
		fwrite($fp,"</testsuite>");
	}
}

//初始化用例xml文件
function init_TC_xml($filename)
{
	$filename = iconv("UTF-8","GB2312",$filename);
	if(!file_exists("./download"))
	{
		mkdir("./download");
	}
	$fp = fopen("./download/$filename","w+");
	$temp = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
	fwrite($fp,$temp);
	return $fp;
}
//读取测试用例表头
function read_TC_header($objSheet,$maxColumn)
{
	$tcinfo = array();
	$currentRow = 1;
	for($currentColumn = 'A';$currentColumn <= "V";$currentColumn ++)
	{
		//echo $currentColumn."</br>";
		$val = htmlspecialchars($objSheet->getCellByColumnAndRow(ord($currentColumn) - 65,$currentRow)->getValue());
		$tcinfo["$val"] = $currentColumn;
	}
	return $tcinfo;
}
//-----需求相关函数--------
//
function init_TR_xml($filename) 
{
	$filename = iconv("UTF-8","GB2312",$filename);
	if(!file_exists("./download"))
	{
		mkdir("./download");
	}
	$fp = fopen("./download/$filename","w+");
	$temp = "<?xml version=\"1.0\" encoding=\"utf-8\"?>
	<requirement-specification>";
	fwrite($fp,$temp);
	return $fp;
}
//
function read_TR_header($objSheet)
{
	$trinfo = array();
	$currentRow = 1;
	for($currentColumn = 'A';$currentColumn <= 'V';$currentColumn++)
	{
		$val = htmlspecialchars($objSheet->getCellByColumnAndRow(ord($currentColumn) - 65,$currentRow)->getValue());
		$trinfo["$val"] = $currentColumn;
	}
	return $trinfo;
}

function template_TR_xml($trinfo,$fp)
{
	$temp = "<requirement>
	<docid><![CDATA[{$trinfo[TR_ID_C]}]]></docid>
	<title><![CDATA[{$trinfo[TR_NAME_C]}]]></title>
	<version>1</version>
	<revision>1</revision>
	<node_order></node_order>
	<description><![CDATA[<p>{$trinfo[TR_D_C]}</p>]]></description>
	<status><![CDATA[F]]></status>
	<type><![CDATA[3]]></type>
	<expected_coverage><![CDATA[{$trinfo[TR_S_C]}]]></expected_coverage>
	<custom_fields>
	<custom_field>
		<name><![CDATA[".REQUIREMENT_I."]]></name>
		<value><![CDATA[{$trinfo[REQUIREMENT_I_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".REQUIREMENT_D."]]></name>
		<value><![CDATA[{$trinfo[REQUIREMENT_D_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".REQUIREMENT_P."]]></name>
		<value><![CDATA[{$trinfo[REQUIREMENT_P_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".REQUIREMENT_F."]]></name>
		<value><![CDATA[{$trinfo[REQUIREMENT_F_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".USER_SCENE."]]></name>
		<value><![CDATA[{$trinfo[USER_SCENE_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".REQUIREMENT_S."]]></name>
		<value><![CDATA[{$trinfo[REQUIREMENT_S_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".REMARK."]]></name>
		<value><![CDATA[{$trinfo[REMARK_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".TR_P."]]></name>
		<value><![CDATA[{$trinfo[TR_P_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".TR_T."]]></name>
		<value><![CDATA[{$trinfo[TR_T_C]}]]></value>
	</custom_field>
	<custom_field>
		<name><![CDATA[".TR_M."]]></name>
		<value><![CDATA[{$trinfo[TR_M_C]}]]></value>
	</custom_field>
	</custom_fields>
	</requirement>";
	fwrite($fp,$temp);
}
//
function template_TR_xml2($trinfo,$fp,$source)
{
	$temp = "<req_spec title=\"{$trinfo[SPECIFICATION_C]}\" doc_id=\"{$trinfo[SPECIFICATION_ID_C]}\">
	<revision><![CDATA[1]]></revision>
	<type><![CDATA[3]]></type>
	<node_order><![CDATA[]]></node_order>
	<total_req><![CDATA[]]></total_req>
	<scope>
	<![CDATA[<p>{$trinfo[SPECIFICATION_C]}</p>]]>
	</scope>
	";
	if($source)
	{
		$temp = "</req_spec>".$temp;
		$source --;
	}
	$source ++;
	fwrite($fp,$temp);
	return $source;
}

function template_TR_end($fp,$source)
{
	for($i = 1;$i <= $source;$i ++)
	{
		fwrite($fp,"</req_spec>");
	}
}

function download_xml($dir,$filename)
{
	$dir = chop($dir);
	$filepath = $dir.$filename;
	$filesize = filesize($filepath);
	header("Content-Type:application/xml");
	header("Accept-Ranges:bytes");
	header("Accept-Length:".$filesize);
	header("Content-Disposition:attachment;filename=".$filename);
	
	$fp = fopen($filepath,"r");
	$buffer_size = 1024;
	$cur_pos = 0;
	
	while(!feof($fp) && $filesize - $cur_pos > $buffer_size)
	{
		$buffer = fread($fp,$buffer_size);
		echo $buffer;
		$cur_pos += $buffer_size;
	}
	
	$buffer = fread($fp,$filesize - $cur_pos);
	echo $buffer;
	fclose($fp);
	return true;
}

function add_tabname_head($fp,$tabname)
{
	$temp = "<req_spec title=\"{$tabname}\" doc_id=\"{$tabname}\">
				<revision><![CDATA[1]]></revision>
					<type><![CDATA[3]]></type>
					<node_order><![CDATA[]]></node_order>
					<total_req><![CDATA[]]></total_req>
					<scope>
					<![CDATA[<p>{$tabname}</p>]]></scope>";
	fwrite($fp,$temp);
}
function add_tabname_end($fp)
{
	$temp = "</req_spec>";
	fwrite($fp,$temp);
}
?>