<html>
<?php
error_reporting(0);
header("Content-Type:text/html;charset=UTF-8");
//error_reporting(E_ERROR | E_PARSE);

require_once 'PHPExcel.php';
require_once './PHPExcel/Writer/Excel5.php';
require_once 'function.php';
require_once 'config.php';

//register_shutdown_function("abend");
$rq_radio = $_POST['tab_radio_rq'];
$filepath = $_FILES['re_xls']['tmp_name'];
if($filepath == "")
{
	echo "上传文件失败！！请重试！";
	echo "<a href='./helpDoc/uploadError.html'>为什么我的文件上传失败？</a>";
}

$objExcel = new PHPExcel();
$objReader = new PHPExcel_Reader_Excel5();
$objExcel = $objReader->load($filepath);
//获取工作表总数和名称
$sheetNames = $objReader->listWorksheetNames($filepath);
//获取工作表总数
$sheetCount = count($sheetNames);

for($currentSheet = 0;$currentSheet <= $sheetCount - 1;$currentSheet ++)
{
	//为每张表创建一个xml文件
	$xmlname = preg_replace("/.xls/i","",$_FILES['re_xls']['name'])."-".$sheetNames[$currentSheet].".xml";
	$xml[$currentSheet] = $xmlname;
	$fp = init_TR_xml($xmlname);
	//初始化值，并创建表对象
	$source = 0;
	$objSheet = $objExcel->getSheet($currentSheet);
	$maxColumn = $objSheet->getHighestColumn();
	$maxRow = $objSheet->getHighestRow();
	$trinfo = read_TR_header($objSheet);
	
	if($rq_radio)
	{
		add_tabname_head($fp,$sheetNames[$currentSheet]);
	}
	
	
	for($currentRow = 2;$currentRow <= $maxRow;$currentRow ++)
	{
		for($currentColumn = 'A';$currentColumn <= 'V';$currentColumn ++)
		{
			$val = htmlspecialchars($objSheet->getCellByColumnAndRow(ord($currentColumn) - 65,$currentRow)->getValue());
			switch ($currentColumn)
			{
				case $trinfo[SPECIFICATION_ID]:
					$trinfo[SPECIFICATION_ID_C] = $val;
					break;
				case $trinfo[SPECIFICATION]:
					$trinfo[SPECIFICATION_C] = $val;
					break;
				case $trinfo[REQUIREMENT_I]://软件需求编号如果为空将向上跟随
					$trinfo[REQUIREMENT_I_C] = $val;
					if(!empty($trinfo[REQUIREMENT_I_C]))
					{
						$trinfo[RE_ID] = $trinfo[REQUIREMENT_I_C];
					}
					else
					{
						$trinfo[REQUIREMENT_I_C] = $trinfo[RE_ID];
					}
					break;
				case $trinfo[REQUIREMENT_D]://软件需求描述如果为空将向上跟随
					$trinfo[REQUIREMENT_D_C] = $val;
					if(!empty($trinfo[REQUIREMENT_D_C]))
					{
						$trinfo[RE_D] = $trinfo[REQUIREMENT_D_C];
					}
					else
					{
						$trinfo[REQUIREMENT_D_C] = $trinfo[RE_D];
					}
					break;
				case $trinfo[REQUIREMENT_P]://软件需求优先级如果为空将向上跟随
					$trinfo[REQUIREMENT_P_C] = $val;
					if(!empty($trinfo[REQUIREMENT_P_C]))
					{
						$trinfo[RE_P] = $trinfo[REQUIREMENT_P_C];
					}
					else
					{
						$trinfo[REQUIREMENT_P_C] = $trinfo[RE_P];
					}
					break;
				case $trinfo[REQUIREMENT_F]:
					$trinfo[REQUIREMENT_F_C] = $val;
					break;
				case $trinfo[USER_SCENE]:
					$trinfo[USER_SCENE_C] = $val;
					break;
				case $trinfo[REQUIREMENT_S]:
					$trinfo[REQUIREMENT_S_C] = $val;
					break;
				case $trinfo[REMARK]:
					$trinfo[REMARK_C] = $val;
					break;
				case $trinfo[TR_ID]:
					$trinfo[TR_ID_C] = $val;
					break;
				case $trinfo[TR_NAME]:
					$trinfo[TR_NAME_C] = $val;
					break;
				case $trinfo[TR_D]:
					$trinfo[TR_D_C] = $val;
					break;
				case $trinfo[TR_P]:
					$trinfo[TR_P_C] = $val;
					break;
				case $trinfo[TR_T]:
					$trinfo[TR_T_C] = $val;
					break;
				case $trinfo[TR_S]:
					$trinfo[TR_S_C] = $val;
					break;
				case $trinfo[TR_M]:
					$trinfo[TR_M_C] = $val;
					break;					
			}
		}
		//当测试需求编号为空时，将跳过该需求的转换--20150216
		//将中止--20150304
		if(empty($trinfo[TR_ID_C]))
		{
			break;
		}
		//echo "【".$trinfo[TR_ID_C]."】";
		if(!empty($trinfo[SPECIFICATION_ID_C]))
		{
			$source = template_TR_xml2($trinfo,$fp,$source);
			template_TR_xml($trinfo,$fp);
		}
		else
		{
			template_TR_xml($trinfo,$fp);
		}
	}
	if($rq_radio)
	{
		add_tabname_end($fp);
	}
	template_TR_end($fp,$source);
	fwrite($fp,'</requirement-specification>');
	fclose($fp);
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