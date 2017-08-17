<html>
<?php
error_reporting(0);
header("Content-Type:text/html;charset=UTF-8");
//error_reporting(E_ERROR | E_WARNING | E_PARSE);

require_once 'PHPExcel.php';
require_once './PHPExcel/Writer/Excel5.php';
require_once 'function.php';
require_once 'config.php';

//register_shutdown_function("abend");

$filepath = $_FILES['xls']['tmp_name'];
if($filepath == "")
{
	echo "上传文件失败！！请重试！";
	echo "<a href='./helpDoc/uploadError.html'>为什么我的文件上传失败？</a>";
}

//获取输入的自定义字段
$cus = get_cus($_POST);
$tab_radio = $_POST["tab_radio"];
//使用PHPExcel获取表对象
$objExcel = new PHPExcel();
$objReader = new PHPExcel_Reader_Excel5();
$objExcel = $objReader->load($filepath);
//获取工作表的总数和名称
$sheetNames = $objReader->listWorksheetNames($filepath);
//获取工作表总数
$sheetCount = count($sheetNames);
//创建XML文件
//$xmlname = preg_replace("/xls/i","",$_FILES['xls']['name'])."xml";
//$fp = init_TC_xml($xmlname);
//遍历表
for($currentSheet = 0;$currentSheet <= $sheetCount - 1;$currentSheet ++)
{
	//为每个表创建相应的xml
	$xmlname = preg_replace("/.xls/i","",$_FILES['xls']['name'])."-".$sheetNames[$currentSheet].".xml";
	$xml[$currentSheet] = $xmlname;
	$fp = init_TC_xml($xmlname);
	//初始化积分，同时创建表操作对象
	$source1 = 0;
	$source2 = 0;
	$objSheet = $objExcel->getSheet($currentSheet);
	//获取单个表中的最大行和最大列 最大列数可能不准
	$maxColumn = $objSheet->getHighestColumn();
	$maxRow = $objSheet->getHighestRow();
	$tcinfo = read_TC_header($objSheet,$maxColumn);
	/*取消对标签页的转化为用例集的功能--20150212
	发现传入空值即可解决问题，so easy！！--20150213
	变成单选框的方式,默认不将标签页作为用例集--20150213*/
	if(!$tab_radio)
	{
		$s = template_xml2("",$fp,0,0);
	}
	elseif($tab_radio)
	{
		$s = template_xml2($sheetNames[$currentSheet],$fp,0,0);
	}
	//遍历行
	for($currentRow = 2;$currentRow <= $maxRow;$currentRow ++)
	{
		//遍历列，从A读到V
		for($currentColumn = 'A';$currentColumn <= 'V';$currentColumn ++)
		{
			$val = htmlspecialchars($objSheet->getCellByColumnAndRow(ord($currentColumn) - 65,$currentRow)->getValue());
			
			switch ($currentColumn)
			{
				//对基本字段的赋值，看起来太长了，可能要做成函数调用
				case $tcinfo[SUITE1]:
					$tcinfo[SUITE1_C] = $val;
					break;
				case $tcinfo[SUITE2]:
					$tcinfo[SUITE2_C] = $val;
					break;
				case $tcinfo[REQUIREMENT]:
					$tcinfo[REQUIREMENT_C] = $val;
					if(!empty($tcinfo[REQUIREMENT_C]))
					{
						$tcinfo[RE_C] = $tcinfo[REQUIREMENT_C];
					}
					else
					{
						$tcinfo[REQUIREMENT_C] = $tcinfo[RE_C];
					}
					break;
				case $tcinfo[TC_NAME]:
					$tcinfo[TC_NAME_C] = $val;
					break;
				//将优先级的“高”“中”“低”转为数字-20150210
				case $tcinfo[PRIORITY]:
					$tcinfo[PRIORITY_C] = $val;
					if($tcinfo[PRIORITY_C] == "高")
					{
						$tcinfo[PRIORITY_C] = 3;
					}
					else if($tcinfo[PRIORITY_C] == "低")
					{
						$tcinfo[PRIORITY_C] = 1;
					}
					else if($tcinfo[PRIORITY_C] == "中")
					{
						$tcinfo[PRIORITY_C] = 2;
					}
					break;
				case $tcinfo[EXECUTION_TYPE]:
					$tcinfo[EXECUTION_TYPE_C] = $val;
					if($tcinfo[EXECUTION_TYPE_C] == "自动的")
					{
						$tcinfo[EXECUTION_TYPE_C] = 2;
					}else{
						$tcinfo[EXECUTION_TYPE_C] = 1;
					}
					break;
				case $tcinfo[KEYWORD]:
					$tcinfo[KEYWORD_C] = $val;
					break;
				case $tcinfo[SUMMARY];
					$tcinfo[SUMMARY_C] = $val;
					if(empty($tcinfo[SUMMARY_C]))
					{
						$tcinfo[SUMMARY_C] = 'NONE';
					}
					break;
				case $tcinfo[PRECONDITION];
					$tcinfo[PRECONDITION_C] = $val;
					if(empty($tcinfo[PRECONDITION_C]))
					{
						$tcinfo[PRECONDITION_C] = 'NONE';
					}
					break;
				case $tcinfo[ACTION]:
					$tcinfo[ACTION_C] = $val;
					break;
				case $tcinfo[EXPECTED_RESULTS]:
					$tcinfo[EXPECTED_RESULTS_C] = $val;
					break;
				case $tcinfo[ESTIMATED_TIME]:
					$tcinfo[ESTIMATED_TIME_C] = $val;
					break;
				case $tcinfo[AUTHOR]:
					$tcinfo[AUTHOR_C] = $val;
					break;
				//对自定义字段的赋值
				case $tcinfo[$cus[1]]:
					$tcinfo['CUS_1'] = $val;
					break;
				case $tcinfo[$cus[2]]:
					$tcinfo['CUS_2'] = $val;
					break;
				case $tcinfo[$cus[3]]:
					$tcinfo['CUS_3'] = $val;
					break;
				case $tcinfo[$cus[4]]:
					$tcinfo['CUS_4'] = $val;
					break;
				case $tcinfo[$cus[5]]:
					$tcinfo['CUS_5'] = $val;
					break;
			}
		}
		/*核心部分，根据单行中包含目录的情况进行写入操作
		  忘了写的时候是怎么想的了！shit！--20150212
		  当测试用例名称为空时，将跳过--20150216
		  将中止--20150304*/
		if(empty($tcinfo[TC_NAME_C]))
		{
			break;
		}
		if(!empty($tcinfo[SUITE1_C]))
		{
			$times = 0;
		}
		if(!empty($tcinfo[SUITE1_C]) && !empty($tcinfo[SUITE2_C]))
		{
			$source1 = template_xml2($tcinfo[SUITE1_C],$fp,$source1,1);
			if(!$times)
			{
				$source1 ++;
				if($source1 >= 3)
				{
					$source1 = 2;
				}
			}
			$times ++;
			$source2 = 0;
			$source2 = template_xml2($tcinfo[SUITE2_C],$fp,$source2,2);
			template_xml($tcinfo,$fp,$cus);
		}
		elseif(!empty($tcinfo[SUITE1_C]) && empty($tcinfo[SUITE2_C]))
		{
			$source1 = template_xml2($tcinfo[SUITE1_C],$fp,$source1,1);
			$source2 = 0;
			template_xml($tcinfo,$fp,$cus);
		}
		elseif(!empty($tcinfo[SUITE2_C]))
		{
			if(!$times)
			{
				$source1 ++;
				if($source1 >=3)
				{
					$source = 2;
				}
			}
			$times ++;
			$source2 = template_xml2($tcinfo[SUITE2_C],$fp,$source2,2);
			template_xml($tcinfo,$fp,$cus);
		}
		else
		{
			template_xml($tcinfo,$fp,$cus);
		}
	}
	//根据之前的遗留<testsuite>个数，写入</testsuite>
	if($source1)
	{
		template_end($fp,$source1);
	}
	else
	{
		template_end($fp,$source2);
	}
	/*为了除去标签页的用例集，在结尾处同时删除一个</testsuite>12字节
	这种方式不好，先凑合用--20150213*/
	//$path = dirname(__FILE__)."\download\\".$xmlname;
	//$path = iconv("utf-8","gb2312//IGNORE",$path);
	//ftruncate($fp,filesize($path) - 12);
	fclose($fp);
	//clearstatcache();
	//unset($path);
	/*本来上面那段代码只有fclose($fp)--20150213;恢复原样--20150213*/
}
echo '<p>转换格式完毕，下载您需要的文件</p>';
for($i = 0;$i <= count($xml) - 1;$i ++)
{
	echo '<form action="download.php" method="post">
	<input type="hidden" name="xmlname" value="'.$xml[$i].'"/>	
	<p>'.$xml[$i].'</p><input type="submit" name="submit" value="下载"/>
	</table>
	</form>';
}
?>
<input type="button" name="back" value="返回" onclick="javascript:history.go(-1)"/>
</html>