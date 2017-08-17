<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=UTF-8"/>
<link rel="StyleSheet" type="text/css" href="./css/style.css"/>
<script>
function checkExName(fileId,formId){	
	var form = document.getElementById(formId);
	var filename = document.getElementById(fileId).value;
	filename = filename.split("\\");//这里要将 \ 转义一下
	var name = filename[filename.length - 1];
	if(name == '')
	{
		alert('请选择文件');
		return;
	}
	pos = name.lastIndexOf(".");
	exname = name.substring(pos,name.length);
	if(!(exname.toLowerCase() == ".xls"))
	{
		alert("上传的文件必须为xls格式");
		return;
	}
	else
	{
		form.submit();
	}
}
</script>
</head>

<body>
<div id="head">
<h2>Testlink导入辅助工具</h2>
</div>
<div id="main">
<div id="requirement">
<form id="tr_form" action="re_process.php" method="post" enctype="multipart/form-data">
	<label for="file">测试需求:</label>
	<input type="file" name="re_xls" id="re_xls"/>
	<br>
	<br>
	<div>
		<input type="radio" name="tab_radio_rq" value=1>标签页作为需求集</input>
		<input type="radio" name="tab_radio_rq" value=0>标签页不作为需求集</input>
	</div>
	<br>
	<input type="button" name="tr_button" value="开始转换" onclick="checkExName('re_xls','tr_form')">
</form>
</div>
<p>____________________________________________________________</p>
<div id="testsuite">
<form id="tc_form" action="process.php" method="post" enctype="multipart/form-data">
	<label for="file">测试用例:</label>
	<input type="file" name="xls" id="xls"/><br>
	<br>
	<div id="ra_tab">
		<input type="radio" name="tab_radio" value=1>标签页作为用例集</input>
		<input type="radio" name="tab_radio" value=0>标签页不作为用例集</input>
	</div>
	<div id="custom">
	<p>输入自定义字段的名称：</p>
		<ul>
			<li>1、<input type="text" name="cus_1" id="cus_1"/></li>
			<li>2、<input type="text" name="cus_2" id="cus_2"/></li>
			<li>3、<input type="text" name="cus_3" id="cus_3"/></li>
			<li>4、<input type="text" name="cus_4" id="cus_4"/></li>
			<li>5、<input type="text" name="cus_5" id="cus_5"/></li>
		</ul>
	</div>	
	<input type="button" name="tc_button" value="开始转换" onclick="checkExName('xls','tc_form')"/>
</form>
</div>
<br>
<a href="./helpDoc/help.html">使用帮助（首次使用必看）</a>
</div>

</body>
</html>
