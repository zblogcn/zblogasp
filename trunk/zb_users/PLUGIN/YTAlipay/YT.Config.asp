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
If CheckPluginState("YTAlipay") = False Then Call ShowError(48)
If CheckPluginState("YTCMS") = False Then
	Response.Write("您没有安装YT.CMS插件,无法进行配置管理")
	Response.End()
End If
Dim Config
Set Config = new YT_Alipay
If Request.Form("action") = "save" Then
	Config.Partner = Request.Form("partner")
	Config.Key = Request.Form("key")
	Config.Seller_Email = Request.Form("seller_email")
	Call Config.Save()
	Response.Redirect("YT.Panel.asp")
End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader">支付宝</div>
  <div class="SubMenu"> <a href="YT.Panel.asp"><span class="m-left m-now">订单管理</span></a><a href="YT.Config.asp"><span class="m-left">系统配置</span></a>
  </div>
  <div id="divMain2">
<form id="form1" name="form1" method="post" action="">
<table width="100%" style="margin-top:0;" cellspacing="0" cellpadding="0" border="0">
<tr>
<td width="12%" align="right">合作者身份ID</td>
<td width="88%"><input type="text" id="partner" name="partner" style="width:95%" value="<%=Config.Partner%>" /></td>
</tr>
<tr>
  <td align="right">安全检验码</td>
  <td><input type="text" id="key" name="key" style="width:95%" value="<%=Config.Key%>" /></td>
</tr>
<tr>
  <td align="right">支付宝帐户</td>
  <td><input type="text" id="seller_email" name="seller_email" style="width:95%" value="<%=Config.Seller_Email%>" /></td>
</tr>
<tr>
  <td align="right">&nbsp;</td>
  <td>
  <input type="hidden" name="action" id="action" value="save" />
    <input type="submit" name="button" id="button" class="button" value="保存设置" />
  </td>
</tr>
</table>
</form>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Set Config = Nothing
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>