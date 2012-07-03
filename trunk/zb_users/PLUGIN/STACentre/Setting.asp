<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 Devo
'// 插件制作:    haphic(http://haphic.com)
'// 备    注:    Deep09 参数设定
'// 最后修改：   2008-2-9
'// 最后版本:    0.4
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function_md5.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("STACentre")=False Then Call ShowError(48)

BlogTitle="静态化中心"


%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../../../ZB_SYSTEM/CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../../../ZB_SYSTEM/script/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="function.js" type="text/javascript"></script>
	<style>
		table{border:none;}
		td{border:1px solid #FFF;}
		a.toggleBtn{display:block;height:24px;width:137px;padding:0 0 0 9px;line-height:12px;color:black;background:url(background.gif) no-repeat 0 0;}
		a.toggleBtn:hover{color:black;text-decoration:none;background:url(background.gif) no-repeat 0 -24px;}
		a.toggleBtn span{display:none;}
		a.toggleBtn span.checked{display:block;height:12px;padding:6px 0 6px 26px;}
		a.toggleBtn span.enable{background:url(background.gif) no-repeat -147px 0;}
		a.toggleBtn span.disable{background:url(background.gif) no-repeat -147px -24px;}
		#globalHintWrapper{position:absolute;left:0;right:0;top:44px;}
		#globalHint{display:none;width:580px;margin:0 auto;color:#333;font-size:18px;line-height:100%;text-align:center;border:1px solid #ffdf9a;}
		#globalHint a{color:#000;font-weight:bold;}
		#globalHint a:hover{color:#333;text-decoration:none;}
		#preview {background:#ffe;}
		#preview a{border:3px solid #dda;}
		#preview a.crrView{border:3px solid #fff;background-color:#ffe;}
		#preview p{padding:5px 10px;background:#dda;}
		#setting a.previewBtn{display:block;width:30px;height:20px;background:url(background.gif) no-repeat 0 -48px;}
		#setting a.previewBtn:hover{background:url(background.gif) no-repeat -30px -48px;}
		#setting a.crrView{background:url(background.gif) no-repeat -30px -48px;}
		#buildStatus p{line-height:100%;padding:8px 0 0 38px;}
	</style>
	<title><%=BlogTitle%></title>
</head>
<body>
			<div id="globalHintWrapper"><p id="globalHint"><span></span> &nbsp; <a href="javascript:void(0);" onclick="$('#globalHint').fadeOut('normal');return false;">[关闭]</a></p></div>
			<div id="divMain">
<div class="Header">静态化中心之参数设定.</div>
<div class="SubMenu">
	<span class="m-left m-now"><a href="Setting.asp">[插件后台管理页]</a> </span>
</div>
<div id="divMain2">
<div id="ShowBlogHint"><%Call GetBlogHint()%></div>

<form id="edit">

<p>
	※ 静态路径配置: 可以是 
	<a href="javascript:void(0);" onclick="InsertText(objTextbox,this.innerHTML,false);checkValue(objTextbox);return false;">{%post%}</a>, 
	<a href="javascript:void(0);" onclick="InsertText(objTextbox,this.innerHTML,false);checkValue(objTextbox);return false;">{%type%}</a>, 
	<a href="javascript:void(0);" onclick="InsertText(objTextbox,this.innerHTML,false);checkValue(objTextbox);return false;">{%alias%}</a>, 
	<a href="javascript:void(0);" onclick="InsertText(objTextbox,this.innerHTML,false);checkValue(objTextbox);return false;">{%id%}</a>, 
	<a href="javascript:void(0);" onclick="InsertText(objTextbox,this.innerHTML,false);checkValue(objTextbox);return false;">{%name%}</a> 
	之间的组合,用 <a href="javascript:void(0);" onclick="InsertText(objTextbox,this.innerHTML,false);checkValue(objTextbox);return false;">/</a> 分隔目录层次. 
	<a href="javascript:makeGlobalHint('help','您可以通过点击这些符号来将其插入文本框!');" title="您可以通过点击这些符号来将其插入文本框!">[?]</a>
</p>
<p>
	※ 推荐以下经典配置: 
	(1) <a href="javascript:void(0);" onclick="FillText(objTextbox,this.innerHTML);checkValue(objTextbox);return false;">{%post%}/{%type%}</a>, 
	(2) <a href="javascript:void(0);" onclick="FillText(objTextbox,this.innerHTML);checkValue(objTextbox);return false;">{%type%}</a>
</p>
<table id="setting">

	<tr>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">静态分类页已启用</span><span class="disable">静态分类页已停用</span></a>
		<input type="hidden" class="checkValue" name="STACentre_Dir_Categorys_Enable" id="STACentre_Dir_Categorys_Enable" value="<%=STACentre_Dir_Categorys_Enable%>"/>
	</td>
	<td><input type="text" class="inputValue" autocomplete="off" style="width:200px;" name="STACentre_Dir_Categorys_Regex" id="STACentre_Dir_Categorys_Regex" value="<%=TransferHTML(STACentre_Dir_Categorys_Regex,"[html-format]")%>" onfocus="objTextbox=this.id;" onkeyup="checkValue(objTextbox);"/></td>
	<td width="20"><a class="previewBtn" title="预览静态路径" href="javascript:void(0);" onclick="return Preview('Categorys');"></a></td>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">匿名路径已启用</span><span class="disable">匿名路径已停用</span></a>
		<input type="hidden" class="checkValue relValue" name="STACentre_Dir_Categorys_Anonymous" id="STACentre_Dir_Categorys_Anonymous" value="<%=STACentre_Dir_Categorys_Anonymous%>"/>
	</td>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">子分类使用子目录</span><span class="disable">所有分类同级目录</span></a>
		<input type="hidden" class="checkValue relValue" name="STACentre_Dir_Categorys_FCate" id="STACentre_Dir_Categorys_FCate" value="<%=STACentre_Dir_Categorys_FCate%>"/>
	</td>
	</tr>
	
	<tr>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">静态Tag页已启用</span><span class="disable">静态Tag页已停用</span></a>
		<input type="hidden" class="checkValue" name="STACentre_Dir_Tags_Enable" id="STACentre_Dir_Tags_Enable" value="<%=STACentre_Dir_Tags_Enable%>"/>
	</td>
	<td><input type="text" class="inputValue" autocomplete="off" style="width:200px;" name="STACentre_Dir_Tags_Regex" id="STACentre_Dir_Tags_Regex" value="<%=TransferHTML(STACentre_Dir_Tags_Regex,"[html-format]")%>" onfocus="objTextbox=this.id;" onkeyup="checkValue(objTextbox);"/></td>
	<td width="20"><a class="previewBtn" title="预览静态路径" href="javascript:void(0);" onclick="return Preview('Tags');"></a></td>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">匿名路径已启用</span><span class="disable">匿名路径已停用</span></a>
		<input type="hidden" class="checkValue relValue" name="STACentre_Dir_Tags_Anonymous" id="STACentre_Dir_Tags_Anonymous" value="<%=STACentre_Dir_Tags_Anonymous%>"/>
	</td>
	<td>&nbsp;</td>
	</tr>

	<tr>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">静态作者页已启用</span><span class="disable">静态作者页已停用</span></a>
		<input type="hidden" class="checkValue" name="STACentre_Dir_Authors_Enable" id="STACentre_Dir_Authors_Enable" value="<%=STACentre_Dir_Authors_Enable%>"/>
	</td>
	<td><input type="text" class="inputValue" autocomplete="off" style="width:200px;" name="STACentre_Dir_Authors_Regex" id="STACentre_Dir_Authors_Regex" value="<%=TransferHTML(STACentre_Dir_Authors_Regex,"[html-format]")%>" onfocus="objTextbox=this.id;" onkeyup="checkValue(objTextbox);"/></td>
	<td width="20"><a class="previewBtn" title="预览静态路径" href="javascript:void(0);" onclick="return Preview('Authors');"></a></td>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">匿名路径已启用</span><span class="disable">匿名路径已停用</span></a>
		<input type="hidden" class="checkValue relValue" name="STACentre_Dir_Authors_Anonymous" id="STACentre_Dir_Authors_Anonymous" value="<%=STACentre_Dir_Authors_Anonymous%>"/>
	</td>
	<td>&nbsp;</td>
	</tr>

	<tr>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">静态归档页已启用</span><span class="disable">静态归档页已停用</span></a>
		<input type="hidden" class="checkValue" name="STACentre_Dir_Archives_Enable" id="STACentre_Dir_Archives_Enable" value="<%=STACentre_Dir_Archives_Enable%>"/>
	</td>
	<td><input type="text" class="inputValue" autocomplete="off" style="width:200px;" name="STACentre_Dir_Archives_Regex" id="STACentre_Dir_Archives_Regex" value="<%=TransferHTML(STACentre_Dir_Archives_Regex,"[html-format]")%>" onfocus="objTextbox=this.id;" onkeyup="checkValue(objTextbox);"/></td>
	<td width="20"><a class="previewBtn" title="预览静态路径" href="javascript:void(0);" onclick="return Preview('Archives');"></a></td>
	<td>
		<a href="javascript:void(0);" onclick="return toggleCheckbox(this);" class="toggleBtn"><span class="enable">匿名路径已启用</span><span class="disable">匿名路径已停用</span></a>
		<input type="hidden" class="checkValue relValue" name="STACentre_Dir_Archives_Anonymous" id="STACentre_Dir_Archives_Anonymous" value="<%=STACentre_Dir_Archives_Anonymous%>"/>
	</td>
	<td>
		格式:
		<select style="width:100px;" class="edit" size="1" onchange="document.getElementById('STACentre_Dir_Archives_Format').value=this.options[this.selectedIndex].value;ValueChanged();">
			<option value="std" <%If STACentre_Dir_Archives_Format="std" Then%>selected="selected"<%End If%>><%=Year(Now)%>-<%=Right("0"&Month(Now),2)%></option>
			<option value="abbr" <%If STACentre_Dir_Archives_Format="abbr" Then%>selected="selected"<%End If%>><%=ZVA_Month_Abbr(Month(Now))%>-<%=Year(Now)%></option>
			<option value="full" <%If STACentre_Dir_Archives_Format="full" Then%>selected="selected"<%End If%>><%=ZVA_Month(Month(Now))%>-<%=Year(Now)%></option>
		</select>
		<input type="hidden" class="inputValue" name="STACentre_Dir_Archives_Format" id="STACentre_Dir_Archives_Format" value="<%=STACentre_Dir_Archives_Format%>"/>
	</td>
	</tr>

</table>

<script language="JavaScript" type="text/javascript">
	getValueSet();
</script>

<hr/>
<div id="preview">
	<p>
		<a href="javascript:void(0);" onclick="return Preview('Categorys');">[预览分类路径]</a> 
		<a href="javascript:void(0);" onclick="return Preview('Tags');">[预览Tag路径]</a>
		<a href="javascript:void(0);" onclick="return Preview('Authors');">[预览作者路径]</a>
		<a href="javascript:void(0);" onclick="return Preview('Archives');">[预览归档路径]</a>
	</p>
	<table>
		<tr><td>请点击菜单来生成预览 ↑</td></tr>
	</table>
</div>
<hr/>

<table id="submit">
	<tr>
		<td>
			<p><input type="submit" class="button" disabled="disabled" value="保存修改" id="btnPost" onclick="return SaveSetting();" /> <input type="button" class="button" value="重建静态列表页" id="btnBuild" onclick="makeGlobalHint('help','← 如进度条长时间停止响应, 请点击左侧 [插件后台管理页] 重新载入本页.');return PageRebuild(0);" /></p><p></p>
		</td>
		<td>
			<div id="buildStatus" style="width:480px;margin:-10px 0 0 10px;padding:0;line-height:100%;"></div>
		</td>
	<tr>
</table>

<script language="JavaScript" type="text/javascript">
	<%	
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(Server.MapPath("progress.txt")) Then
				Response.Write "$('#setting,#preview').hide();"
				Response.Write "$('#buildStatus').html('<img src=\""point.gif\"" style=\""float:left\""/><p>继续上次未完成的重建!</p>');"
			End If
		Set fso=Nothing
	%>
</script>

</form>
</div>
</div>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>

