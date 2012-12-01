<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    33195@qq.com
'// 程序名称:    	YT.Build
'// 开始时间:    	2011.03.26
'// 最后修改:    2012.08.24
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
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
If CheckPluginState("YTBuild")=False Then Call ShowError(48)

If Request.Form("action") = "save" Then
	Dim k,o,a
		On Error Resume Next
		For Each a In BlogConfig.Meta.Names
			If a<>"ZC_BLOG_VERSION" Then
				Call Execute("Call BlogConfig.Write("""&a&""","&a&")")
			End If
		Next
		Err.Clear
	Set o = New TConfig
		o.Load "YTBuild"
		For Each k In Request.Form
			If exists(k)<>False Then
				Call BlogConfig.Write(k,Request.Form(k))
			End If
			If exists(k,true)<>False Then
				o.Write k,Request.Form(k)
			End If
		Next
		o.Save
	Set o = Nothing
	Call SaveConfig2Option()
	Call SetBlogHint_Custom("设置已保存!")
	Response.Redirect("YT.Config.asp")
End If
%>
<script language="javascript" type="text/javascript" runat="server">
	var exists=function(s,b){return b?(/BUILD_[A-Z]+/.test(s)?s:false):(/ZC_[A-Z]+_(REGEX|MODE)/.test(s)?s:false);};
</script>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script language="javascript" src="Script/YT.Build.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
$(document).ready(function(){
	var e;
	$('#ZC_STATIC_MODE').change(function(){
		var ZC=YT.Config($(this).val());
		$('select,input').each(function(){
			if(/ZC_[A-Z]+_REGEX/.test($(this).attr('id'))){
				$(this).val(eval('ZC.'+$(this).attr('id')));
			}	
		});							 
	});
	$('input[type=text]').each(function(){
		$(this).focus(function(){e=this;});
	});
	$('em label').each(function(){
		$(this).click(function(){
			YT.Insert($(e)[0],$(this).text());				   
		});					  
	});
});
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader">YT.Build</div>
  <div class="SubMenu"> <a href="YT.Panel.asp"><span class="m-left">控制面板</span></a><a href="YT.Config.asp"><span class="m-left m-now">系统配置</span></a>
  </div>
  <div id="divMain2">
  <form id="form1" name="form1" method="post" action="">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="14%" align="right">开启</td>
    <td width="86%">
      <label>
        <select name="ZC_STATIC_MODE" id="ZC_STATIC_MODE">
          <option value="REWRITE"<%If ZC_STATIC_MODE="REWRITE" Then%> selected="selected"<%End If%>>是</option>
          <option value="ACTIVE"<%If ZC_STATIC_MODE="ACTIVE" Then%> selected="selected"<%End If%>>否</option>
        </select>
      </label>
    </td>
  </tr>
  <tr>
    <td align="right">参数（点击）</td>
    <td>
      <em>
        <label>{%host%}</label>
        <label>{%post%}</label>
        <label>{%category%}</label>
        <label>{%user%}</label>
        <label>{%year%}</label>
        <label>{%month%}</label>
        <label>{%day%}</label>
        <label>{%id%}</label>
        <label>{%alias%}</label>
      </em>
    </td>
  </tr>
  <tr>
    <td align="right">首页</td>
    <td>
<input id="ZC_DEFAULT_REGEX" name="ZC_DEFAULT_REGEX" style="width:500px;" type="text" value="<%=ZC_DEFAULT_REGEX%>" /></td>
  </tr>
  <tr>
    <td align="right">分类页</td>
    <td><input id="ZC_CATEGORY_REGEX" name="ZC_CATEGORY_REGEX" style="width:500px;" type="text" value="<%=ZC_CATEGORY_REGEX%>" /></td>
  </tr>
  <tr>
    <td align="right">作者页</td>
    <td><input id="ZC_USER_REGEX" name="ZC_USER_REGEX" style="width:500px;" type="text" value="<%=ZC_USER_REGEX%>" /></td>
  </tr>
  <tr>
    <td align="right">TAGS页</td>
    <td><input id="ZC_TAGS_REGEX" name="ZC_TAGS_REGEX" style="width:500px;" type="text" value="<%=ZC_TAGS_REGEX%>" /></td>
  </tr>
  <tr>
    <td align="right">日期页</td>
    <td><input id="ZC_DATE_REGEX" name="ZC_DATE_REGEX" style="width:500px;" type="text" value="<%=ZC_DATE_REGEX%>" /></td>
  </tr>
  <tr>
    <td align="right">文章页</td>
    <td><input id="ZC_ARTICLE_REGEX" name="ZC_ARTICLE_REGEX" style="width:500px;" type="text" value="<%=ZC_ARTICLE_REGEX%>" /></td>
  </tr>
  <tr>
    <td align="right">单页</td>
    <td><input id="ZC_PAGE_REGEX" name="ZC_PAGE_REGEX" style="width:500px;" type="text" value="<%=ZC_PAGE_REGEX%>" /></td>
  </tr>
  <tr>
  <td align="right"></td>
    <td></td>
  </tr>
	<%
	Dim oTConfig
    Set oTConfig = new TConfig
		With oTConfig
		.Load "YTBuild"
    %>  
  <tr>
    <td align="right">（发布|编辑）文章（重建）</td>
    <td><table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><input id="BUILD_HOME" name="BUILD_HOME" type="text" value="<%=.Read("BUILD_HOME")%>" class="checkbox" />首页</td>
    <td><input id="BUILD_CATE" name="BUILD_CATE" type="text" value="<%=.Read("BUILD_CATE")%>" class="checkbox" />分类</td>
    <td><input id="BUILD_TAG" name="BUILD_TAG" type="text" value="<%=.Read("BUILD_TAG")%>" class="checkbox" />TAG</td>
    <td><input id="BUILD_USER" name="BUILD_USER" type="text" value="<%=.Read("BUILD_USER")%>" class="checkbox" />作者</td>
    <td><input id="BUILD_DATE" name="BUILD_DATE" type="text" value="<%=.Read("BUILD_DATE")%>" class="checkbox" />归档</td>
  </tr>
</table>
    </td>
  </tr>
  <%
  	End With
  Set oTConfig = Nothing
  %>
  <tr>
    <td align="right"></td>
    <td>
    <input type="hidden" name="action" id="action" value="save" />
    <label>
      <input type="submit" name="button" id="button" class="button" value="保存" />
    </label></td>
  </tr>
</table>
</form>
</div>
</div>
<script type="text/javascript">ActiveLeftMenu("aYTBuildMng");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->