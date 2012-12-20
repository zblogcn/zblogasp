<%
'///////////////////////////////////////////////////////////////////////////////
'// 月上之木 2012.8.25
'///////////////////////////////////////////////////////////////////////////////

Call Add_Response_Plugin("Response_Plugin_Html_Js_Add",vbCrlf &"$.getScript(""" & BlogHost & "zb_users/PLUGIN/SweetTitles/jquery.sweetTitles.js"");")
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add",vbCrlf &"$(""head"").append(""<link rel='stylesheet' type='text/css' href='" & BlogHost & "zb_users/PLUGIN/SweetTitles/sweetTitles.css'>"");")

'注册插件
Call RegisterPlugin("SweetTitles","")

%>