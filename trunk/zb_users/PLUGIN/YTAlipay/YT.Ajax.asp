<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    13120003225@qq.com
'// 程序名称:    	Content Manage System
'// 开始时间:    	2011.03.26
'// 最后修改:    
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
If CheckPluginState("YTCMS") = False Then Call ShowError(48)
Dim Action
	Action = Request("Action")
	Select Case Action
		Case "GetJsonList":Response.Write(new YT_Alipay.GetJsonList(Request("intPage")))
		Case "DelAlipayOrder":Call new YT_Alipay.DelAlipayOrder(Request.QueryString("ID"))
		Case "Delivery":Response.Write(Join(new YT_Alipay.Send_goods_confirm_by_platform(Array("service=send_goods_confirm_by_platform","partner="&new YT_Alipay.Partner,"trade_no="&request.Form("trade_no"),"logistics_name="&request.Form("logistics_name"),"invoice_no="&request.Form("invoice_no"),"transport_type="&request.Form("transport_type"),"_input_charset="&new YT_Alipay.Input_Charset), Split(jsonToObject(request.Form("node")).join(","),",")),"|"))
	End Select
Call System_Terminate()
If Err.Number<>0 then
  Call ShowError(0)
End If
%>