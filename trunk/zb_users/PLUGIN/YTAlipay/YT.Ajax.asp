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
		Case "Delivery":
			Dim s,t
			Set t = new YT_Alipay
			s = Join(new YT_Alipay.Send_goods_confirm_by_platform(Array("service=send_goods_confirm_by_platform","partner="&t.sPartner,"trade_no="&request.Form("trade_no"),"logistics_name="&request.Form("logistics_name"),"invoice_no="&request.Form("invoice_no"),"transport_type="&request.Form("transport_type"),"_input_charset="&t.sInput_Charset), Split(YT.eval(request.Form("node")).join(","),",")),"|")
			Dim Sql
			Dim out_trade_no,trade_no,Rs,YT_ALIPAY_TYPE
			out_trade_no = Request.Form("out_trade_no")
			trade_no = Request.Form("trade_no")
			YT_ALIPAY_TYPE = Request.Form("Service")
			Sql = "UPDATE [YT_Alipay] SET [trade_no] = '"&trade_no&"',[Status] = "&t.GetStatus(t.sStatus(YT_ALIPAY_TYPE),Split(s,"|")(0))&" WHERE [OrderID] = '"&out_trade_no&"'"
			objConn.Execute(Sql)
			Response.Write(s)
			Set Rs = Nothing
			Set t = Nothing
	End Select
Call System_Terminate()
If Err.Number<>0 then
  Call ShowError(0)
End If
%>