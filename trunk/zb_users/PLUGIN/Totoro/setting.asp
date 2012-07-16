<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%
'On Error Resume Next
 %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("Totoro")=False Then Call ShowError(48)

BlogTitle="TotoroⅢ（基于TotoroⅡ的Z-Blog的评论管理审核系统增强版）"

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../../../ZB_SYSTEM/CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../../../ZB_SYSTEM/script/common.js" type="text/javascript"></script>
	<title><%=BlogTitle%></title>
</head>
<body>

			<div id="divMain">
<div class="Header"><%=BlogTitle%></div>

<div class="SubMenu"><span class="m-left m-now"><a href="setting.asp">TotoroⅢ设置</a></span><span class="m-left"><a href="setting1.asp">审核评论<%
	Dim objRS1
	Set objRS1=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [log_ID]<0")
	If (Not objRS1.bof) And (Not objRS1.eof) Then
		Response.Write "("&objRS1(0)&"条未审核的评论)"
	End If
%></a></span></div>

<div id="divMain2">
<form id="edit" name="edit" method="post">
<%

	Call Totoro_Initialize

	Response.Write "<p><b>关于TotoroⅢ</b></p>"
	Response.Write "<p>Totoro是个采用评分机制的防止垃圾留言的插件，原作<a href=""http://www.rainbowsoft.org/"" target=""_blank"">zx.asd</a>。<br/>TotoroⅡ是<a href=""http://ZxMYS.COM"" target=""_blank"">Zx.MYS</a>在Totoro的基础上修改而成的增强版，加入了诸多新特性，同时修正一些问题。<br/>TotoroⅢ是由<a href=""http://www.zsxsoft.com"" target=""_blank"">ZSXSOFT</a>将TotoroII升级到1.9版本后增添新特性的版本。</p>"
	Response.Write "<p>Spam Value(SV)初始值为0，经过相关运算后的SV分值越高Spam嫌疑越大，超过设定的阈值这条评论就进入审核状态。</p>"
	

	Response.Write "<p></p>"
	Response.Write "<p><b>加分减分细则：</b></p>"
	
	Dim strZC_TOTORO_HYPERLINK_VALUE
	strZC_TOTORO_HYPERLINK_VALUE=Totoro_Config.Read("TOTORO_HYPERLINK_VALUE")
	strZC_TOTORO_HYPERLINK_VALUE=TransferHTML(strZC_TOTORO_HYPERLINK_VALUE,"[html-format]")
	Response.Write "<p>1.评论里有链接就加<input name=""strZC_TOTORO_HYPERLINK_VALUE"" style=""width:25px"" type=""text"" value=""" & strZC_TOTORO_HYPERLINK_VALUE & """/>分(默认：10)，每多一个链接SV翻倍加分</p>"
	
	Dim strTOTORO_INTERVAL_VALUE
	strTOTORO_INTERVAL_VALUE=Totoro_Config.Read("TOTORO_INTERVAL_VALUE")
	strTOTORO_INTERVAL_VALUE=TransferHTML(strTOTORO_INTERVAL_VALUE,"[html-format]")
	Response.Write "<p>2.提交频率评分:基数为<input name=""strZC_TOTORO_INTERVAL_VALUE"" style=""width:25px"" type=""text"" value=""" & strTOTORO_INTERVAL_VALUE & """/>分(默认：25)，根据1小时内同一IP的评论数量加分。(每条评论最多加基数的5/6，最少加基数的1/5，按时间间隔递减。)</p>"
	
	Dim strTOTORO_BADWORD_VALUE
	strTOTORO_BADWORD_VALUE=Totoro_Config.Read("TOTORO_BADWORD_VALUE")
	strTOTORO_BADWORD_VALUE=TransferHTML(strTOTORO_BADWORD_VALUE,"[html-format]")
	Response.Write "<p>3.评论里的每一个黑词都加<input name=""strZC_TOTORO_BADWORD_VALUE"" style=""width:25px"" type=""text"" value=""" & strTOTORO_BADWORD_VALUE & """/>分(默认：50)</p>"
	
	Dim strTOTORO_LEVEL_VALUE
	strTOTORO_LEVEL_VALUE=Totoro_Config.Read("TOTORO_LEVEL_VALUE")
	strTOTORO_LEVEL_VALUE=TransferHTML(strTOTORO_LEVEL_VALUE,"[html-format]")
	Response.Write "<p>4.用户信任度评分:基数为<input name=""strZC_TOTORO_LEVEL_VALUE"" style=""width:25px"" type=""text"" value=""" & strTOTORO_LEVEL_VALUE & """/>分(默认：100)，初级用户评论时SV减基数×1，中级用户SV减基数×2，高级用户减SV减基数×3，管理员SV减基数×4</p>"

	Dim strTOTORO_NAME_VALUE
	strTOTORO_NAME_VALUE=Totoro_Config.Read("TOTORO_NAME_VALUE")
	strTOTORO_NAME_VALUE=TransferHTML(strTOTORO_NAME_VALUE,"[html-format]")
	Response.Write "<p>5.访客熟悉度评分:基数为<input name=""strZC_TOTORO_NAME_VALUE"" style=""width:25px"" type=""text"" value=""" & strTOTORO_NAME_VALUE & """/>分(默认：45)，同一访客在BLOG留言1-10条内的SV减10分,10-20条的SV减10分再减基数×1，20-50条的SV减10分再减基数×2，大于50条的SV减10分再减基数×3</p>"
	
	Dim strTOTORO_NUMBER_VALUE
	strTOTORO_NUMBER_VALUE=Totoro_Config.Read("TOTORO_NUMBER_VALUE")
	strTOTORO_NUMBER_VALUE=TransferHTML(strTOTORO_NUMBER_VALUE,"[html-format]")
	Response.Write "<p>6.数字长度评分:基数为<input name=""strTOTORO_NUMBER_VALUE"" style=""width:25px"" type=""text"" value=""" & strTOTORO_NUMBER_VALUE & """/>分(默认：10)。若数字长度达到10位，自动加上基数。多几位，加几次基数。"
	
	
	Dim strZC_TOTORO_SV_THRESHOLD
	strZC_TOTORO_SV_THRESHOLD=Totoro_Config.Read("TOTORO_SV_THRESHOLD")
	strZC_TOTORO_SV_THRESHOLD=TransferHTML(strZC_TOTORO_SV_THRESHOLD,"[html-format]")
	Response.Write "<p>·设置系统审核阈值(默认50，阈值越小越严格，低于0则使游客的评论全进入审核):</p><p><input name=""strZC_TOTORO_SV_THRESHOLD"" style=""width:99%"" type=""text"" value=""" & strZC_TOTORO_SV_THRESHOLD & """/></p><p></p>"

	Dim strZC_TOTORO_BADWORD_LIST
	strZC_TOTORO_BADWORD_LIST=Totoro_Config.Read("TOTORO_BADWORD_LIST")
		strZC_TOTORO_BADWORD_LIST=TransferHTML(strZC_TOTORO_BADWORD_LIST,"[html-format]")
		Response.Write "<p>·黑词列表(分隔符'|'):</p><p><textarea rows=""6"" name=""strZC_TOTORO_BADWORD_LIST"" style=""width:99%"" >"& strZC_TOTORO_BADWORD_LIST &"</textarea></p>"	
	Response.Write "<p>·"

	Dim bolTOTORO_ConHuoxingwen
	bolTOTORO_ConHuoxingwen=Totoro_Config.Read("TOTORO_ConHuoxingwen")
	Response.Write "<input name=""bolTOTORO_ConHuoxingwen"" id=""bolTOTORO_ConHuoxingwen"" type=""checkbox"" value=""True"""
	
	If bolTOTORO_ConHuoxingwen="True" then
		Response.Write " checked=""checked""/>"
	else
		Response.Write "/>"
	End if
	Response.Write "<label for=""bolTOTORO_ConHuoxingwen"">自动转换火星文（将把希腊字母、俄文字母、罗马数字、列表符、全角字母数字标点、汉语拼音、?转换为半角英文字母、半角数字、半角符号再进行下一步操作，不影响实际显示的评论）</label></p><p></p>"	

	Response.Write "<p>·"
	Dim bolTOTORO_DEL_DIRECTLY
	bolTOTORO_DEL_DIRECTLY=Totoro_Config.Read("TOTORO_DEL_DIRECTLY")
	Response.Write "<input name=""bolTOTORO_DEL_DIRECTLY"" id=""bolTOTORO_DEL_DIRECTLY"" type=""checkbox"" value=""True"""
	If bolTOTORO_DEL_DIRECTLY="True" then
		Response.Write " checked=""checked"">"
	else
		Response.Write ">"
	End if

	Response.Write "<label for=""bolTOTORO_DEL_DIRECTLY"">点击[这是SPAM]按钮提取域名后直接删除评论（若不删除则进入审核）</label></p><p></p>"
	Response.Write "<hr/>"
	Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='document.getElementById(""edit"").action=""savesetting.asp"";' /></p>"

	
	'Response.Write "<br/><p><a target='_blank' href='http://bbs.rainbowsoft.org/viewthread.php?tid=11849'>Totoro的相关说明文档</a></p><br/>"


%>
</form>
</div>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
