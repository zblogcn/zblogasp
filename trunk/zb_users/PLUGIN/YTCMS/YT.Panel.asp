<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    33195@qq.com
'// 程序名称:    	Content Manage System
'// 开始时间:    	2011.03.26
'// 最后修改:    2012-08-08
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #INCLUDE FILE="../../C_OPTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_FUNCTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_LIB.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_BASE.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_EVENT.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_PLUGIN.ASP" -->
<!-- #INCLUDE FILE="../../PLUGIN/P_CONFIG.ASP" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
If CheckPluginState("YTCMS") = False Then Call ShowError(48)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
<meta name="generator" content="Z-Blog <%=ZC_BLOG_VERSION%>" />
<meta name="robots" content="nofollow" />
<title><%=ZC_BLOG_TITLE & ZC_MSG044 & BlogTitle%></title>
<link href="<%=BlogHost%>ZB_SYSTEM/CSS/admin2.css" rel="stylesheet" type="text/css" media="screen">
<link rel="stylesheet" rev="stylesheet" href="STYLE/YT.Style.css" type="text/css" media="screen">
<script src="<%=BlogHost%>ZB_SYSTEM/script/common.js" type="text/javascript"></script>
<script src="<%=BlogHost%>ZB_SYSTEM/function/c_admin_js_add.asp" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
	var ZC_BLOG_HOST='<%=ZC_BLOG_HOST%>';
	var ZC_BLOG_THEME = '<%=ZC_BLOG_THEME%>';
	var YT_CMS_XML_URL = ZC_BLOG_HOST+'ZB_USERS/THEME/'+ZC_BLOG_THEME+'/';
</script>
<script language="javascript" src="Config.js" type="text/javascript"></script>
<script language="javascript" src="SCRIPT/YT.Lib.js" type="text/javascript"></script>
<script language="javascript" src="SCRIPT/YT.Interface.js" type="text/javascript"></script>
<script language="javascript" src="SCRIPT/YT.Main.js" type="text/javascript"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div class="wrap d_wrap">
	<div class="wrapHeader">
	<h2>Content Manage System
		<span class="d_themedesc">当前版本：<span id="version"></span> &nbsp;&nbsp; 作者：<a target="_blank" href="" id="author"></a> &nbsp;&nbsp;E-Mail：<a href="#" id="email"></a> &nbsp;&nbsp;<a target="_blank" href="" id="bug">问题反馈</a>&nbsp;&nbsp;<a target="_blank" href="" id="give">愛心赞助</a></span>
	</h2>
    </div>
	
	<form method="post">
		<div class="d_tab"><a class="d_tab_on" href="javascript:void(0)">模块</a><a href="javascript:void(0)">模型</a><a href="javascript:void(0)">模板</a><a href="javascript:void(0)">导入</a><a href="javascript:void(0)">演示</a></div>
					<div id="block" class="d_mainbox" style="display: block;">
				<div class="d_desc"><input type="button" value="新建模块" class="button-primary">调用频繁的代码可使用模块,支持标签及语法,需要显示多行文本框新增模块内容时请回车</div>
				<ul class="d_inner">
					<li class="d_li">
			</li><li class="d_li" style="display:none;">
			<h4></h4>
			<span class="d_check"><label><input type="checkbox">保存</label><label><input type="checkbox">删除</label></span>
			</li></ul>
			<div class="d_desc d_desc_b"><input type="button" value="保存设置" class="button-primary"></div>
			</div>
			
					<div id="model" class="d_mainbox" style="display: none;">
				<div class="d_desc"><input type="button" value="新建模型" class="button-primary">创建表及字段,如不绑定栏目则显示为系统表,绑定多个栏目请按SHIFT+左键;CTRL+左键,<font color=red>注意:字段首字母不能為h字母</font></div>
				<ul class="d_inner">
					<li class="d_li">		
			</li>
            <li class="d_li" style="display:none;">
			<h4>产品中心</h4>
            <span class="d_check"><label><input type="checkbox">修改</label><label><input type="checkbox">删除</label></span>
            <div class="d_adviewcon" style="display:none;">
            <ul>
            	<li>
                表名<input type="text" value="YT_Product" />
                描述<input type="text" value="产品中心" />
                <input type="button" value="新增字段" />
                </li>
                <li>
                    字段<input type="text" value="ID" />
                    描述<input type="text" value="主键" />
                    UI默认值<input type="text" value="主键" />
                    属性<select>
                    <option value="VARCHAR">文本</option>
                    <option value="TEXT">备注</option>
                    <option value="INT">数字</option>
                    <option value="DATETIME">时间</option>
                    <option value="COUNTER(1,1)">自动编号</option>
                    </select>
                    显示UI <select>
                    <option value="text">单行文本框</option>
                    <option value="checkbox">多选</option>
                    <option value="select">下拉框</option>
                    <option value="textarea">多行文本框</option>
                    <option value="upload-image">上传图片</option>
                    <option value="upload-attachment">上传附件</option>
                    </select>
                    <em style="display:none">×</em>
                </li>
                <li>
                    <select multiple="multiple" size="5" style="width:660px;">
                    <%
					dim aryCateInOrder,m,n
                       aryCateInOrder=GetCategoryOrder()
                        For m=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
                            If Categorys(aryCateInOrder(m)).ParentID=0 Then
                                Response.Write "<option value="""&Categorys(aryCateInOrder(m)).ID&""" "
                                Response.Write ">"&TransferHTML( Categorys(aryCateInOrder(m)).Name,"[html-format]")&"</option>"
                    
                                For n=0 To UBound(aryCateInOrder)
                                    If Categorys(aryCateInOrder(n)).ParentID=Categorys(aryCateInOrder(m)).ID Then
                                        Response.Write "<option value="""&Categorys(aryCateInOrder(n)).ID&""" "
                                        Response.Write ">&nbsp;└ "&TransferHTML( Categorys(aryCateInOrder(n)).Name,"[html-format]")&"</option>"
                                    End If
                                Next
                            End If
                        Next
                    %>
                    </select>
                </li>
            </ul>
            </div>
            <div class="d_status"></div>
			</li>
            </ul>
			<div class="d_desc d_desc_b"><input type="button" value="保存设置" class="button-primary"></div>
			</div>
			<div id="tpl" class="d_mainbox" style="display: none;">
				<div class="d_desc"><input type="button" value="新建模板" class="button-primary">制作主题,管理当前主题目录的HTML文件,新建模板时不需要输入文件后缀</div>
				<ul class="d_inner">
					<li class="d_li">		
			</li><li class="d_li" style="display:none;">
			<h4></h4>
			<span class="d_check"><label><input type="checkbox">保存</label><label><input type="checkbox">删除</label></span>
			<div class="d_adviewcon"></div>
			<textarea type="textarea" class="d_tarea"></textarea>	
			</li></ul>
			<div class="d_desc d_desc_b"><input type="button" value="保存设置" class="button-primary"></div>
			</div>
            <div id="sql" class="d_mainbox" style="display: none;">
				<div class="d_desc">ACCESS数据导入</div>
				<ul class="d_inner">
					<li class="d_li">	
			</li><li class="d_li" style="display:none;">
			<span class="d_check">
            	<label><input type="checkbox"></label>
            </span>
			</li></ul>
			<div class="d_desc d_desc_b"><input type="button" value="开始导入" class="button-primary"></div>
			</div>
<div id="demo" class="d_mainbox" style="display: none;">	
				<div class="d_desc">常用标签及语法演示</div>
				<ul class="d_inner">
					<li class="d_li">
			</li><li class="d_li">
			
			</li></ul>
			</div>	
		<input type="hidden" value="save" name="action">
	</form>
</div>
<script type="text/javascript">ActiveLeftMenu("aYTCMSMng");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>